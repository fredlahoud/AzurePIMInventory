#Requires -Version 5.1
<#
    AzPimInventory module.

    Inventories Azure PIM role assignments (eligible and active) beneath a management
    group using Azure Resource Manager only. Shared by the command-line script
    (Export-AzPimAssignmentInventory.ps1) and the desktop dashboard (App\).
#>

Set-StrictMode -Version Latest

$script:ArmApiVersion = '2020-10-01'
$script:MgApiVersion  = '2020-05-01'

# Column order of the CSV. Kept explicit so the file layout is a stable contract.
$script:ReportColumns = @(
    'PrincipalName', 'PrincipalType', 'PrincipalEmail', 'PrincipalId',
    'RoleName', 'RoleType', 'PrivilegedRole', 'AssignmentKind', 'MemberType',
    'SubscriptionName', 'ScopeName', 'ScopeType', 'Scope', 'ResourceUri',
    'StartDateUtc', 'ExpiryDateUtc', 'DaysRemaining', 'Status', 'Condition', 'InstanceId'
)

# Roles the Azure portal lists as privileged administrator roles.
$script:DefaultPrivilegedRoleName = @(
    'Owner'
    'Contributor'
    'User Access Administrator'
    'Role Based Access Control Administrator'
)

#region Helpers ---------------------------------------------------------------

function Write-Step {
    param([string]$Message)
    Write-Host "  $Message" -ForegroundColor Cyan
}

function Assert-Prerequisite {
    if (-not (Get-Module -ListAvailable -Name Az.Accounts)) {
        throw "The Az.Accounts module is required. Install it with: Install-Module Az.Accounts -Scope CurrentUser"
    }
    Import-Module Az.Accounts -ErrorAction Stop

    $context = Get-AzContext -ErrorAction SilentlyContinue
    if (-not $context) {
        throw "No Azure context found. Run Connect-AzAccount (or Connect-AzAccount -Identity in Azure Automation) first."
    }
    Write-Step "Signed in as $($context.Account.Id) on tenant $($context.Tenant.Id)"
}

function Get-SafeProperty {
    param($Object, [string]$Name, $Default = $null)

    if ($null -eq $Object) { return $Default }
    if ($Object.PSObject.Properties.Name -notcontains $Name) { return $Default }
    $value = $Object.$Name
    if ($null -eq $value) { return $Default }
    return $value
}

<#
    Works out how long to wait before retrying a throttled or failed ARM call.
    Honours Retry-After in either of its RFC 9110 forms (delta-seconds or an
    HTTP-date), and otherwise falls back to exponential backoff. Never throws.
#>
function Get-RetryDelay {
    param($Response, [int]$Attempt = 1)

    $fallback = [int][math]::Min(60, [math]::Pow(2, [math]::Max(1, $Attempt)))
    $headers  = Get-SafeProperty -Object $Response -Name 'Headers'
    if (-not $headers) { return $fallback }

    $value = $null
    try {
        foreach ($h in @($headers.GetEnumerator())) {
            if ([string]$h.Key -eq 'Retry-After') { $value = [string](@($h.Value)[0]); break }
        }
    }
    catch { return $fallback }
    if ([string]::IsNullOrWhiteSpace($value)) { return $fallback }

    $seconds = 0
    if ([int]::TryParse($value.Trim(), [ref]$seconds)) {
        return [int][math]::Max(1, [math]::Min(300, $seconds))
    }

    $when = [DateTimeOffset]::MinValue
    if ([DateTimeOffset]::TryParse($value, [cultureinfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::AssumeUniversal, [ref]$when)) {
        $delta = [math]::Ceiling(($when - [DateTimeOffset]::UtcNow).TotalSeconds)
        return [int][math]::Max(1, [math]::Min(300, $delta))
    }

    return $fallback
}

<#
    Invokes an ARM GET, following nextLink paging and retrying on throttling.
    Returns the concatenated 'value' arrays. Throws on hard failure so the caller
    can decide whether a single scope failure is fatal.
#>
function Invoke-ArmList {
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$MaxRetry = 5
    )

    $results = [System.Collections.Generic.List[object]]::new()
    $next    = $Path

    while ($next) {
        $attempt  = 0
        $response = $null

        while ($true) {
            $attempt++
            try {
                $response = Invoke-AzRestMethod -Path $next -Method GET -ErrorAction Stop
            }
            catch {
                if ($attempt -ge $MaxRetry) { throw }
                Start-Sleep -Seconds (Get-RetryDelay -Response $null -Attempt $attempt)
                continue
            }

            # 429 and 5xx are transient. Honour Retry-After when present.
            if ($response.StatusCode -eq 429 -or $response.StatusCode -ge 500) {
                if ($attempt -ge $MaxRetry) {
                    throw "ARM returned $($response.StatusCode) after $attempt attempts for $next"
                }
                Start-Sleep -Seconds (Get-RetryDelay -Response $response -Attempt $attempt)
                continue
            }

            break
        }

        if ($response.StatusCode -ne 200) {
            throw "ARM returned $($response.StatusCode) for $next : $($response.Content)"
        }

        $payload = $response.Content | ConvertFrom-Json
        $value   = Get-SafeProperty -Object $payload -Name 'value'
        if ($value) {
            foreach ($item in $value) { $results.Add($item) }
        }

        $next     = $null
        $nextLink = Get-SafeProperty -Object $payload -Name 'nextLink'
        if ($nextLink) {
            # nextLink is an absolute URL; Invoke-AzRestMethod -Path wants a relative path.
            $next = $nextLink -replace '^https://[^/]+', ''
        }
    }

    return , $results.ToArray()
}

<#
    Returns the root management group plus every descendant management group and
    subscription, as a list of ARM scope strings.
#>
function Get-TargetScope {
    param(
        [Parameter(Mandatory)][string]$ManagementGroupId,
        [hashtable]$SubscriptionNames = @{}
    )

    $rootScope = "/providers/Microsoft.Management/managementGroups/$ManagementGroupId"
    $scopes    = [System.Collections.Generic.List[string]]::new()
    $scopes.Add($rootScope)

    $path        = "$rootScope/descendants?api-version=$script:MgApiVersion"
    $descendants = Invoke-ArmList -Path $path

    foreach ($d in $descendants) {
        # Descendants are either nested management groups or subscriptions.
        if ($d.type -like '*managementGroups/subscriptions') {
            $scopes.Add("/subscriptions/$($d.name)")
            $properties  = Get-SafeProperty -Object $d -Name 'properties'
            $displayName = Get-SafeProperty -Object $properties -Name 'displayName' -Default ''
            if (-not [string]::IsNullOrWhiteSpace($displayName)) {
                $SubscriptionNames[$d.name] = $displayName
            }
        }
        elseif ($d.type -like '*managementGroups') {
            $scopes.Add("/providers/Microsoft.Management/managementGroups/$($d.name)")
        }
    }

    return , @($scopes | Select-Object -Unique)
}

function Get-PimInstance {
    param(
        [Parameter(Mandatory)][string]$Scope,
        [Parameter(Mandatory)][ValidateSet('Eligible', 'Active')][string]$Kind,
        [switch]$IncludeFuture,
        [AllowNull()][System.Collections.Generic.List[object]]$ErrorLog
    )

    # Schedule instances are assignments active now. Schedules also include
    # assignments scheduled to start in the future.
    $resource = switch ($Kind) {
        'Eligible' { if ($IncludeFuture) { 'roleEligibilitySchedules' } else { 'roleEligibilityScheduleInstances' } }
        'Active'   { if ($IncludeFuture) { 'roleAssignmentSchedules' } else { 'roleAssignmentScheduleInstances' } }
    }
    $path = "$Scope/providers/Microsoft.Authorization/$resource" + "?api-version=$script:ArmApiVersion"

    try {
        $items = Invoke-ArmList -Path $path
    }
    catch {
        # A single unreadable scope must not abort the whole inventory. Record and continue.
        if ($null -ne $ErrorLog) {
            $ErrorLog.Add([pscustomobject]@{
                    Scope   = $Scope
                    Kind    = $Kind
                    Message = $_.Exception.Message
                })
        }
        Write-Warning "Could not read $Kind assignments at $Scope : $($_.Exception.Message)"
        return , @()
    }

    foreach ($i in $items) {
        Add-Member -InputObject $i -NotePropertyName '_kind' -NotePropertyValue $Kind -Force
    }
    return , $items
}

<#
    The same assignment is returned by several scope queries because ARM reports
    at, above and below each scope. Keeps the first instance of each id.
#>
function Get-UniqueInstance {
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Instance)

    $seen   = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $unique = [System.Collections.Generic.List[object]]::new()
    foreach ($i in $Instance) {
        $id = [string](Get-SafeProperty -Object $i -Name 'id' -Default '')
        if ([string]::IsNullOrEmpty($id) -or $seen.Add($id)) { $unique.Add($i) }
    }
    return , $unique.ToArray()
}

#endregion

#region Transform -------------------------------------------------------------

function Get-ScopeType {
    param([string]$Scope)

    if ([string]::IsNullOrWhiteSpace($Scope)) { return 'Unknown' }
    switch -Regex ($Scope) {
        '^/$'                                                      { return 'Tenant root' }
        '/providers/Microsoft\.Management/managementGroups/[^/]+$' { return 'Management group' }
        '^/subscriptions/[^/]+$'                                   { return 'Subscription' }
        '^/subscriptions/[^/]+/resourceGroups/[^/]+$'              { return 'Resource group' }
        '^/subscriptions/[^/]+/resourceGroups/[^/]+/providers/.+'  { return 'Resource' }
        default                                                    { return 'Other' }
    }
}

<#
    Normalises an ARM timestamp to a UTC DateTime. Handles ISO 8601 strings and the
    DateTime objects that PowerShell 7's ConvertFrom-Json produces from them.
    Returns $null for an empty value or one that cannot be parsed.
#>
function ConvertTo-UtcDateTime {
    param($Value)

    if ($null -eq $Value) { return $null }
    if ($Value -is [DateTimeOffset]) { return $Value.UtcDateTime }
    if ($Value -is [datetime]) {
        if ($Value.Kind -eq [DateTimeKind]::Unspecified) {
            return [datetime]::SpecifyKind($Value, [DateTimeKind]::Utc)
        }
        return $Value.ToUniversalTime()
    }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }

    $parsed = [DateTimeOffset]::MinValue
    if ([DateTimeOffset]::TryParse($text, [cultureinfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::AssumeUniversal, [ref]$parsed)) {
        return $parsed.UtcDateTime
    }
    return $null
}

<#
    Flattens raw ARM instances into report rows and computes expiry state.
    Kept separate from the ARM calls so it can be exercised with sample data.
#>
function ConvertTo-ReportRow {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Instance,
        [Parameter(Mandatory)][int]$WarningDays,
        [Parameter(Mandatory)][datetime]$Now,
        [hashtable]$SubscriptionNames = @{},
        [string[]]$PrivilegedRoleName = @()
    )

    $rows  = [System.Collections.Generic.List[object]]::new()
    $names = $SubscriptionNames.Clone()

    # Resolve offline subscription names before processing any child-scope rows.
    foreach ($i in $Instance) {
        $p         = Get-SafeProperty -Object $i -Name 'properties'
        $scopeInfo = Get-SafeProperty -Object (Get-SafeProperty -Object $p -Name 'expandedProperties') -Name 'scope'
        $scope     = Get-SafeProperty -Object $p -Name 'scope' -Default (Get-SafeProperty -Object $scopeInfo -Name 'id' -Default '')
        if ($scope -match '^/subscriptions/([^/]+)$') {
            $subscriptionId = $Matches[1]
            $displayName    = Get-SafeProperty -Object $scopeInfo -Name 'displayName' -Default ''
            if (-not [string]::IsNullOrWhiteSpace($displayName) -and
                (-not $names.ContainsKey($subscriptionId) -or
                 [string]::IsNullOrWhiteSpace($names[$subscriptionId]))) {
                $names[$subscriptionId] = $displayName
            }
        }
    }

    foreach ($i in $Instance) {
        $p        = Get-SafeProperty -Object $i -Name 'properties'
        $expanded = Get-SafeProperty -Object $p -Name 'expandedProperties'

        $principal      = Get-SafeProperty -Object $expanded -Name 'principal'
        $roleDefinition = Get-SafeProperty -Object $expanded -Name 'roleDefinition'
        $scopeInfo      = Get-SafeProperty -Object $expanded -Name 'scope'

        $scope = Get-SafeProperty -Object $p -Name 'scope' -Default (Get-SafeProperty -Object $scopeInfo -Name 'id' -Default '')
        $subscriptionName = ''
        if ($scope -match '^/subscriptions/([^/]+)(?:/|$)') {
            $subscriptionId = $Matches[1]
            if (-not $names.ContainsKey($subscriptionId) -or
                [string]::IsNullOrWhiteSpace($names[$subscriptionId])) {
                $names[$subscriptionId] = "(unresolved: $subscriptionId)"
                Write-Warning "Subscription name unavailable for $subscriptionId in hierarchy metadata or assignment data."
            }
            $subscriptionName = $names[$subscriptionId]
        }

        $endRaw   = Get-SafeProperty -Object $p -Name 'endDateTime'
        $startRaw = Get-SafeProperty -Object $p -Name 'startDateTime'
        $end      = ConvertTo-UtcDateTime $endRaw
        $start    = ConvertTo-UtcDateTime $startRaw

        # An end date that is present but unreadable must never be reported as
        # 'Permanent'; that would hide exactly the kind of row this report exists for.
        $endUnreadable = ($null -eq $end) -and -not [string]::IsNullOrWhiteSpace([string]$endRaw)
        if ($endUnreadable) {
            Write-Warning "Unrecognised endDateTime '$endRaw' on $(Get-SafeProperty -Object $i -Name 'id' -Default '(no id)')."
        }

        $daysRemaining = $null
        if ($end) { $daysRemaining = [int][math]::Floor(($end - $Now).TotalDays) }

        $status = if ($endUnreadable) { 'Unknown' }
        elseif ($null -ne $daysRemaining -and $daysRemaining -lt 0) { 'Expired' }
        elseif ($start -and $start -gt $Now) { 'Pending' }
        elseif (-not $end) { 'Permanent' }
        elseif ($daysRemaining -le $WarningDays) { 'Expiring' }
        else { 'Active' }

        $kind           = Get-SafeProperty -Object $i -Name '_kind' -Default 'Unknown'
        $assignmentType = Get-SafeProperty -Object $p -Name 'assignmentType' -Default ''

        # Eligible instances have no assignmentType. Active ones are either standing
        # access ('Assigned') or a live just-in-time activation ('Activated').
        $category = if ($kind -eq 'Eligible') { 'Eligible' }
        elseif ($assignmentType -eq 'Activated') { 'Activation (JIT)' }
        else { 'Active' }

        $roleName = Get-SafeProperty -Object $roleDefinition -Name 'displayName' -Default '(unknown role)'

        $rows.Add([pscustomobject]@{
                PrincipalName    = Get-SafeProperty -Object $principal -Name 'displayName' -Default '(unresolved)'
                PrincipalType    = Get-SafeProperty -Object $p -Name 'principalType' -Default (Get-SafeProperty -Object $principal -Name 'type' -Default '')
                PrincipalEmail   = Get-SafeProperty -Object $principal -Name 'email' -Default ''
                PrincipalId      = Get-SafeProperty -Object $p -Name 'principalId' -Default ''
                RoleName         = $roleName
                RoleType         = Get-SafeProperty -Object $roleDefinition -Name 'type' -Default ''
                PrivilegedRole   = if ($PrivilegedRoleName -contains $roleName) { 'Yes' } else { 'No' }
                AssignmentKind   = $category
                MemberType       = Get-SafeProperty -Object $p -Name 'memberType' -Default ''
                SubscriptionName = $subscriptionName
                ScopeName        = Get-SafeProperty -Object $scopeInfo -Name 'displayName' -Default ''
                ScopeType        = Get-ScopeType $scope
                Scope            = $scope
                ResourceUri      = $scope
                StartDateUtc     = if ($start) { $start.ToString('yyyy-MM-dd HH:mm', [cultureinfo]::InvariantCulture) } else { '' }
                ExpiryDateUtc    = if ($end) { $end.ToString('yyyy-MM-dd HH:mm', [cultureinfo]::InvariantCulture) } else { '' }
                DaysRemaining    = if ($null -ne $daysRemaining) { $daysRemaining } else { '' }
                Status           = $status
                Condition        = if (Get-SafeProperty -Object $p -Name 'condition' -Default '') { 'Yes' } else { 'No' }
                InstanceId       = Get-SafeProperty -Object $i -Name 'id' -Default ''
            })
    }

    return , $rows.ToArray()
}

<#
    True when $Scope is one of the in-hierarchy scopes, or sits beneath one of the
    in-hierarchy subscriptions. Exact matching, not string prefixes: a prefix test
    would treat '.../managementGroups/mg-prod-legacy' as inside '.../mg-prod'.
#>
function Test-ScopeInHierarchy {
    param(
        [AllowEmptyString()][string]$Scope,
        [Parameter(Mandatory)][System.Collections.Generic.HashSet[string]]$HierarchyScope
    )

    if ([string]::IsNullOrEmpty($Scope)) { return $false }
    $normalised = $Scope.TrimEnd('/')
    if ($HierarchyScope.Contains($normalised)) { return $true }
    if ($normalised -match '^(/subscriptions/[^/]+)(?:/|$)') {
        return $HierarchyScope.Contains($Matches[1])
    }
    return $false
}

function Select-InHierarchyRow {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Row,
        [Parameter(Mandatory)][string[]]$Scope
    )

    $set = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($s in $Scope) { [void]$set.Add($s.TrimEnd('/')) }
    return , @($Row | Where-Object { Test-ScopeInHierarchy -Scope $_.Scope -HierarchyScope $set })
}

#endregion

#region Output ----------------------------------------------------------------

<#
    Neutralises spreadsheet formula injection (CWE-1236). Display names are set by
    users, guests and application owners, so a name such as '=HYPERLINK(...)' would
    otherwise execute when the CSV is opened in Excel.
#>
function Protect-CsvValue {
    param([AllowNull()][AllowEmptyString()][string]$Value)

    if ([string]::IsNullOrEmpty($Value)) { return $Value }
    if ($Value -match '^[=+\-@\t\r]') { return "'" + $Value }
    return $Value
}

function Export-ReportCsv {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Row,
        [Parameter(Mandatory)][string]$Path
    )

    $textColumns = @('PrincipalName', 'PrincipalEmail', 'RoleName', 'SubscriptionName', 'ScopeName')
    $safe = foreach ($r in $Row) {
        $copy = [ordered]@{}
        foreach ($column in $script:ReportColumns) {
            $value = Get-SafeProperty -Object $r -Name $column -Default ''
            $copy[$column] = if ($textColumns -contains $column) { Protect-CsvValue $value } else { $value }
        }
        [pscustomobject]$copy
    }

    $lines = if (@($safe).Count -gt 0) {
        @($safe | ConvertTo-Csv -NoTypeInformation)
    }
    else {
        @(($script:ReportColumns | ForEach-Object { '"' + $_ + '"' }) -join ',')
    }

    # UTF-8 with BOM on every PowerShell version, so Excel renders non-ASCII names.
    [System.IO.File]::WriteAllLines($Path, [string[]]$lines, [System.Text.UTF8Encoding]::new($true))
}

function ConvertTo-HtmlText {
    param($Value)
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function ConvertTo-HtmlReport {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Row,
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][int]$WarningDays,
        [datetime]$Now = [datetime]::UtcNow,
        [AllowEmptyCollection()][object[]]$ScopeError = @()
    )

    $total      = $Row.Count
    $expired    = @($Row | Where-Object { $_.Status -eq 'Expired' }).Count
    $expiring   = @($Row | Where-Object { $_.Status -eq 'Expiring' }).Count
    $permanent  = @($Row | Where-Object { $_.Status -eq 'Permanent' }).Count
    $pending    = @($Row | Where-Object { $_.Status -eq 'Pending' }).Count
    $unknown    = @($Row | Where-Object { $_.Status -eq 'Unknown' }).Count
    $eligible   = @($Row | Where-Object { $_.AssignmentKind -eq 'Eligible' }).Count
    $active     = @($Row | Where-Object { $_.AssignmentKind -eq 'Active' }).Count
    $standing   = @($Row | Where-Object { $_.AssignmentKind -eq 'Active' -and $_.PrivilegedRole -eq 'Yes' }).Count

    $sb = [System.Text.StringBuilder]::new()

    # Sort so the rows that need action appear first.
    # Test for a blank with IsNullOrEmpty, not -eq '': PowerShell coerces the empty
    # string to 0, so "0 -eq ''" is true and an assignment expiring today would
    # otherwise sort to the bottom alongside the permanent ones.
    $order  = @{ 'Unknown' = 0; 'Expired' = 1; 'Expiring' = 2; 'Pending' = 3; 'Active' = 4; 'Permanent' = 5 }
    $sorted = $Row | Sort-Object `
        @{ Expression = { $order[[string]$_.Status] } }, `
        @{ Expression = { if ([string]::IsNullOrEmpty([string]$_.DaysRemaining)) { [int]::MaxValue } else { [int]$_.DaysRemaining } } }

    foreach ($r in $sorted) {
        $cls = switch ($r.Status) {
            'Expired'   { 'st-expired' }
            'Expiring'  { 'st-expiring' }
            'Permanent' { 'st-permanent' }
            'Pending'   { 'st-pending' }
            'Unknown'   { 'st-expired' }
            default     { 'st-ok' }
        }
        $priv      = if ($r.PrivilegedRole -eq 'Yes') { '1' } else { '0' }
        $daysSort  = if ([string]::IsNullOrEmpty([string]$r.DaysRemaining)) { '999999' } else { [string]$r.DaysRemaining }
        $dateSort  = if ([string]::IsNullOrEmpty([string]$r.ExpiryDateUtc)) { '9999' } else { [string]$r.ExpiryDateUtc }

        [void]$sb.Append("<tr data-status=""$(ConvertTo-HtmlText $r.Status)"" data-kind=""$(ConvertTo-HtmlText $r.AssignmentKind)"" data-priv=""$priv"">")
        [void]$sb.Append("<td>$(ConvertTo-HtmlText $r.PrincipalName)</td>")
        [void]$sb.Append("<td>$(ConvertTo-HtmlText $r.PrincipalType)</td>")
        [void]$sb.Append("<td>$(ConvertTo-HtmlText $r.RoleName)</td>")
        [void]$sb.Append("<td>$(if ($priv -eq '1') { '<span class="pill st-priv">Privileged</span>' })</td>")
        [void]$sb.Append("<td>$(ConvertTo-HtmlText $r.AssignmentKind)</td>")
        [void]$sb.Append("<td>$(ConvertTo-HtmlText $r.SubscriptionName)</td>")
        [void]$sb.Append("<td>$(ConvertTo-HtmlText $r.ScopeType)</td>")
        [void]$sb.Append("<td title=""$(ConvertTo-HtmlText $r.Scope)"">$(ConvertTo-HtmlText $r.ScopeName)</td>")
        [void]$sb.Append("<td class=""resource-uri"">$(ConvertTo-HtmlText $r.ResourceUri)</td>")
        [void]$sb.Append("<td data-sort=""$(ConvertTo-HtmlText $dateSort)"">$(ConvertTo-HtmlText $r.ExpiryDateUtc)</td>")
        [void]$sb.Append("<td class=""num"" data-sort=""$(ConvertTo-HtmlText $daysSort)"">$(ConvertTo-HtmlText $r.DaysRemaining)</td>")
        [void]$sb.Append("<td><span class=""pill $cls"">$(ConvertTo-HtmlText $r.Status)</span></td>")
        [void]$sb.Append('</tr>')
    }

    $errorBlock = ''
    if ($ScopeError.Count -gt 0) {
        $items = ($ScopeError | ForEach-Object {
                "<li>$(ConvertTo-HtmlText $_.Scope) ($(ConvertTo-HtmlText $_.Kind)): $(ConvertTo-HtmlText $_.Message)</li>"
            }) -join ''
        $errorBlock = "<div class=""warn""><strong>$($ScopeError.Count) scope(s) could not be read.</strong> The inventory below is incomplete.<ul>$items</ul></div>"
    }

    $optionalCards = ''
    if ($pending -gt 0) { $optionalCards += "<div class=""card pend""><div class=""n"">$pending</div><div class=""l"">Pending (future start)</div></div>" }
    if ($unknown -gt 0) { $optionalCards += "<div class=""card bad""><div class=""n"">$unknown</div><div class=""l"">Unreadable expiry</div></div>" }

    $generated = $Now.ToString('dd MMM yyyy HH:mm', [cultureinfo]::InvariantCulture)
    $rowsHtml  = $sb.ToString()
    $titleHtml = ConvertTo-HtmlText $Title

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>$titleHtml</title>
<style>
 :root { --ink:#1b1b1f; --muted:#5f6368; --line:#e3e3e6; --bg:#faf9f8; --panel:#fff; --head:#f3f2f1; --hover:#f7f9fc; }
 @media (prefers-color-scheme: dark) {
   :root { --ink:#e6e6e6; --muted:#a0a0a0; --line:#3a3a3e; --bg:#1b1b1f; --panel:#242428; --head:#2d2d32; --hover:#2a2f38; }
 }
 * { box-sizing:border-box; }
 body { font-family:'Segoe UI',system-ui,-apple-system,sans-serif; margin:0; background:var(--bg); color:var(--ink); }
 header { background:#0f6cbd; color:#fff; padding:20px 28px; }
 header h1 { margin:0; font-size:19px; font-weight:600; }
 header p { margin:4px 0 0; font-size:13px; opacity:.9; }
 main { padding:22px 28px 60px; }
 .cards { display:flex; flex-wrap:wrap; gap:12px; margin-bottom:20px; }
 .card { background:var(--panel); border:1px solid var(--line); border-radius:8px; padding:14px 18px; min-width:132px; }
 .card .n { font-size:26px; font-weight:600; line-height:1.1; }
 .card .l { font-size:12px; color:var(--muted); margin-top:3px; }
 .card.bad .n { color:#d13438; } .card.warnc .n { color:#c47a00; } .card.perm .n { color:#8e3bc4; } .card.pend .n { color:#0f6cbd; }
 .card.priv { border-color:#d13438; }
 .warn { background:#fff4ce; color:#3b2e00; border:1px solid #e9c96a; border-radius:8px; padding:12px 16px; margin-bottom:18px; font-size:13px; }
 .warn ul { margin:8px 0 0 18px; }
 .toolbar { display:flex; gap:10px; margin-bottom:12px; flex-wrap:wrap; align-items:center; }
 input[type=search], select, button { padding:8px 11px; border:1px solid var(--line); border-radius:6px; font-size:13px; font-family:inherit; background:var(--panel); color:var(--ink); }
 button { cursor:pointer; }
 button:hover { background:var(--head); }
 label { font-size:13px; display:flex; gap:6px; align-items:center; }
 input[type=search] { min-width:290px; }
 .count { font-size:12px; color:var(--muted); margin-left:auto; }
 .table-wrap { overflow-x:auto; }
 table { width:100%; border-collapse:collapse; background:var(--panel); border:1px solid var(--line); border-radius:8px; overflow:hidden; font-size:13px; }
 th { background:var(--head); text-align:left; padding:10px 12px; font-weight:600; cursor:pointer; white-space:nowrap; user-select:none; border-bottom:1px solid var(--line); }
 th::after { content:' \2195'; opacity:.35; font-size:11px; }
 td { padding:9px 12px; border-bottom:1px solid var(--line); vertical-align:top; }
 td.num { text-align:right; font-variant-numeric:tabular-nums; }
 td.resource-uri { min-width:240px; overflow-wrap:anywhere; color:var(--muted); }
 tbody tr:hover { background:var(--hover); }
 .pill { display:inline-block; padding:2px 9px; border-radius:11px; font-size:11.5px; font-weight:600; white-space:nowrap; }
 .st-expired { background:#fde7e9; color:#a4262c; }
 .st-expiring { background:#fff4ce; color:#8a5700; }
 .st-ok { background:#dff6dd; color:#0b6a0b; }
 .st-permanent { background:#f3e8fb; color:#7719aa; }
 .st-pending { background:#deecf9; color:#0f548c; }
 .st-priv { background:#fde7e9; color:#a4262c; }
 footer { margin-top:18px; font-size:12px; color:var(--muted); }
 .hidden { display:none; }
</style>
</head>
<body>
<header>
  <h1>$titleHtml</h1>
  <p>Data as of $generated UTC &nbsp;|&nbsp; Expiry warning threshold: $WarningDays days &nbsp;|&nbsp; Source: Azure Resource Manager (no Microsoft Graph)</p>
</header>
<main>
$errorBlock
<div class="cards">
  <div class="card"><div class="n">$total</div><div class="l">Total assignments</div></div>
  <div class="card bad"><div class="n">$expired</div><div class="l">Expired</div></div>
  <div class="card warnc"><div class="n">$expiring</div><div class="l">Expiring &le; $WarningDays days</div></div>
  <div class="card perm"><div class="n">$permanent</div><div class="l">No expiry</div></div>
  <div class="card"><div class="n">$eligible</div><div class="l">Eligible</div></div>
  <div class="card"><div class="n">$active</div><div class="l">Active</div></div>
  <div class="card bad priv"><div class="n">$standing</div><div class="l">Standing privileged access</div></div>
  $optionalCards
</div>

<div class="toolbar">
  <input type="search" id="q" placeholder="Filter by principal, role, subscription, scope..." aria-label="Filter rows">
  <select id="statusFilter" aria-label="Status">
    <option value="">All statuses</option>
    <option value="Expired">Expired</option>
    <option value="Expiring">Expiring</option>
    <option value="Pending">Pending</option>
    <option value="Active">Active</option>
    <option value="Permanent">Permanent</option>
    <option value="Unknown">Unknown</option>
  </select>
  <select id="kindFilter" aria-label="Assignment kind">
    <option value="">All assignment kinds</option>
    <option value="Eligible">Eligible</option>
    <option value="Active">Active</option>
    <option value="Activation (JIT)">Activation (JIT)</option>
  </select>
  <label><input type="checkbox" id="privFilter"> Privileged roles only</label>
  <button type="button" id="exportCsv">Export visible rows</button>
  <span class="count" id="count"></span>
</div>

<div class="table-wrap">
<table id="t">
<thead><tr>
  <th>Principal</th><th>Type</th><th>Role</th><th>Privileged</th><th>Assignment</th>
  <th>Subscription</th><th>Scope type</th><th>Scope</th><th>Resource URI</th>
  <th>Expiry (UTC)</th><th>Days left</th><th>Status</th>
</tr></thead>
<tbody>
$rowsHtml
</tbody>
</table>
</div>
<footer>Hover a scope name to see its full Azure resource path. Click any column heading to sort. "Standing privileged access" counts active, non-JIT assignments of a privileged role.</footer>
</main>
<script>
(function () {
  var table = document.getElementById('t');
  var body = table.tBodies[0];
  var rows = Array.prototype.slice.call(body.rows);
  var q = document.getElementById('q');
  var sf = document.getElementById('statusFilter');
  var kf = document.getElementById('kindFilter');
  var pf = document.getElementById('privFilter');
  var count = document.getElementById('count');

  function applyFilter() {
    var term = q.value.toLowerCase();
    var visible = 0;
    rows.forEach(function (r) {
      var ok = (!term || r.textContent.toLowerCase().indexOf(term) !== -1) &&
               (!sf.value || r.getAttribute('data-status') === sf.value) &&
               (!kf.value || r.getAttribute('data-kind') === kf.value) &&
               (!pf.checked || r.getAttribute('data-priv') === '1');
      r.classList.toggle('hidden', !ok);
      if (ok) { visible++; }
    });
    count.textContent = 'Showing ' + visible + ' of ' + rows.length;
  }
  q.addEventListener('input', applyFilter);
  sf.addEventListener('change', applyFilter);
  kf.addEventListener('change', applyFilter);
  pf.addEventListener('change', applyFilter);
  applyFilter();

  function sortValue(cell) {
    var v = cell.getAttribute('data-sort');
    return v === null ? cell.textContent.trim() : v;
  }

  var dir = {};
  Array.prototype.forEach.call(table.tHead.rows[0].cells, function (th, idx) {
    th.addEventListener('click', function () {
      dir[idx] = !dir[idx];
      var sign = dir[idx] ? 1 : -1;
      rows.sort(function (a, b) {
        var x = sortValue(a.cells[idx]);
        var y = sortValue(b.cells[idx]);
        var nx = Number(x), ny = Number(y);
        if (x !== '' && y !== '' && !isNaN(nx) && !isNaN(ny)) { return (nx - ny) * sign; }
        return x.localeCompare(y) * sign;
      });
      rows.forEach(function (r) { body.appendChild(r); });
    });
  });

  function csvCell(text) {
    var s = text.replace(/\s+/g, ' ').trim();
    if (/^[=+\-@]/.test(s) && isNaN(Number(s))) { s = "'" + s; }
    return '"' + s.replace(/"/g, '""') + '"';
  }
  document.getElementById('exportCsv').addEventListener('click', function () {
    var lines = [Array.prototype.map.call(table.tHead.rows[0].cells, function (c) { return csvCell(c.textContent); }).join(',')];
    rows.forEach(function (r) {
      if (r.classList.contains('hidden')) { return; }
      lines.push(Array.prototype.map.call(r.cells, function (c) { return csvCell(c.textContent); }).join(','));
    });
    var blob = new Blob(['\ufeff' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'pim-inventory-filtered.csv';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  });
})();
</script>
</body>
</html>
"@

    return $html
}

#endregion


#region Public ----------------------------------------------------------------

function Write-InventoryProgress {
    param([hashtable]$State, [string]$Message, [int]$Percent = -1)

    if ($null -eq $State) { return }
    $State['Message'] = $Message
    $State['Percent'] = $Percent
}

<#
.SYNOPSIS
    Collects the PIM assignment inventory and returns it as an object.

.DESCRIPTION
    Live mode enumerates the management group hierarchy and queries every scope.
    Offline mode reads a JSON file of sample instances. The returned object carries
    the report rows plus the metadata needed to render or export them.

.PARAMETER ProgressState
    Optional synchronised hashtable. 'Message' and 'Percent' are updated as the
    collection runs, so a UI on another thread can display progress.
#>
function Get-PimInventory {
    [CmdletBinding(DefaultParameterSetName = 'Live')]
    param(
        [Parameter(Mandatory, ParameterSetName = 'Live')]
        [ValidateNotNullOrEmpty()]
        [string]$RootManagementGroupId,

        [Parameter(Mandatory, ParameterSetName = 'Offline')]
        [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
        [string]$OfflineSampleData,

        [Parameter(ParameterSetName = 'Offline')]
        [datetime]$ReferenceDateUtc,

        [ValidateRange(1, 365)]
        [int]$ExpiryWarningDays = 30,

        [ValidateNotNull()]
        [string[]]$PrivilegedRoleName = $script:DefaultPrivilegedRoleName,

        [switch]$IncludeActivations,

        [switch]$IncludeAncestorScopes,

        [switch]$IncludeFutureAssignments,

        [hashtable]$ProgressState
    )

    $now = if ($PSBoundParameters.ContainsKey('ReferenceDateUtc')) { ConvertTo-UtcDateTime $ReferenceDateUtc } else { [datetime]::UtcNow }
    $instances         = [System.Collections.Generic.List[object]]::new()
    $scopeErrors       = [System.Collections.Generic.List[object]]::new()
    $subscriptionNames = @{}
    $scopes            = @()
    $mgCount           = 0
    $subCount          = 0

    if ($PSCmdlet.ParameterSetName -eq 'Offline') {
        Write-InventoryProgress -State $ProgressState -Message 'Reading sample data'
        Write-Step "Offline preview using $OfflineSampleData"
        $raw = Get-Content -LiteralPath $OfflineSampleData -Raw | ConvertFrom-Json
        foreach ($i in $raw) { $instances.Add($i) }
        $title = 'Azure PIM assignment inventory (sample data)'
    }
    else {
        Assert-Prerequisite

        Write-InventoryProgress -State $ProgressState -Message "Enumerating hierarchy beneath '$RootManagementGroupId'"
        Write-Step "Enumerating hierarchy beneath '$RootManagementGroupId'"
        $scopes   = Get-TargetScope -ManagementGroupId $RootManagementGroupId -SubscriptionNames $subscriptionNames
        $mgCount  = @($scopes | Where-Object { $_ -like '*managementGroups*' }).Count
        $subCount = @($scopes | Where-Object { $_ -like '/subscriptions/*' }).Count
        Write-Step "Found $mgCount management group(s) and $subCount subscription(s)"
        if ($IncludeFutureAssignments) { Write-Step 'Including assignments scheduled to start in the future' }

        $n = 0
        foreach ($scope in $scopes) {
            $n++
            $percent = [int](($n / $scopes.Count) * 100)
            Write-InventoryProgress -State $ProgressState -Message "Reading scope $n of $($scopes.Count)" -Percent $percent
            Write-Progress -Activity 'Collecting PIM assignments' -Status $scope -PercentComplete $percent
            foreach ($kind in @('Eligible', 'Active')) {
                $found = Get-PimInstance -Scope $scope -Kind $kind -IncludeFuture:$IncludeFutureAssignments -ErrorLog $scopeErrors
                foreach ($item in $found) { $instances.Add($item) }
            }
        }
        Write-Progress -Activity 'Collecting PIM assignments' -Completed

        $title = "Azure PIM assignment inventory - $RootManagementGroupId"
    }

    Write-InventoryProgress -State $ProgressState -Message 'Building report'
    $unique = Get-UniqueInstance -Instance $instances.ToArray()
    Write-Step "Collected $($instances.Count) record(s), $($unique.Count) unique after de-duplication"

    $rows = ConvertTo-ReportRow -Instance $unique -WarningDays $ExpiryWarningDays -Now $now `
        -SubscriptionNames $subscriptionNames -PrivilegedRoleName $PrivilegedRoleName

    $excludedActivations = 0
    if (-not $IncludeActivations) {
        $before = $rows.Count
        $rows   = @($rows | Where-Object { $_.AssignmentKind -ne 'Activation (JIT)' })
        $excludedActivations = $before - $rows.Count
        if ($excludedActivations -gt 0) {
            Write-Step "Excluded $excludedActivations transient just-in-time activation(s). Use -IncludeActivations to keep them."
        }
    }

    $excludedAncestors = 0
    if (-not $IncludeAncestorScopes -and $PSCmdlet.ParameterSetName -eq 'Live') {
        $before = $rows.Count
        $rows   = Select-InHierarchyRow -Row $rows -Scope $scopes
        $excludedAncestors = $before - $rows.Count
        if ($excludedAncestors -gt 0) {
            Write-Step "Excluded $excludedAncestors assignment(s) inherited from above the root. Use -IncludeAncestorScopes to keep them."
        }
    }

    Write-InventoryProgress -State $ProgressState -Message 'Done' -Percent 100

    [pscustomobject]@{
        PSTypeName            = 'AzPimInventory.Result'
        Title                 = $title
        Source                = $PSCmdlet.ParameterSetName
        RootManagementGroupId = if ($PSCmdlet.ParameterSetName -eq 'Live') { $RootManagementGroupId } else { '' }
        GeneratedUtc          = $now
        ExpiryWarningDays     = $ExpiryWarningDays
        Rows                  = @($rows)
        ScopeErrors           = $scopeErrors.ToArray()
        ManagementGroupCount  = $mgCount
        SubscriptionCount     = $subCount
        RecordCount           = $instances.Count
        UniqueCount           = $unique.Count
        ExcludedActivations   = $excludedActivations
        ExcludedAncestors     = $excludedAncestors
    }
}

<#
.SYNOPSIS
    Lists the management groups the signed-in identity can see, tenant root first.
#>
function Get-PimManagementGroup {
    [CmdletBinding()]
    param()

    $items  = Invoke-ArmList -Path "/providers/Microsoft.Management/managementGroups?api-version=$script:MgApiVersion"
    $groups = foreach ($m in $items) {
        $properties = Get-SafeProperty -Object $m -Name 'properties'
        $name       = [string](Get-SafeProperty -Object $m -Name 'name' -Default '')
        if ([string]::IsNullOrEmpty($name)) { continue }
        $display    = [string](Get-SafeProperty -Object $properties -Name 'displayName' -Default $name)
        $tenantId   = [string](Get-SafeProperty -Object $properties -Name 'tenantId' -Default '')
        [pscustomobject]@{
            PSTypeName   = 'AzPimInventory.ManagementGroup'
            Id           = $name
            DisplayName  = $display
            TenantId     = $tenantId
            IsTenantRoot = (-not [string]::IsNullOrEmpty($tenantId)) -and $name -eq $tenantId
            Label        = if ($display -eq $name) { $name } else { "$display ($name)" }
        }
    }
    return , @($groups | Sort-Object @{ Expression = { -not $_.IsTenantRoot } }, DisplayName)
}

<#
.SYNOPSIS
    Writes report rows to CSV (UTF-8 with BOM, formula-injection safe).
#>
function Export-PimInventoryCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Row,
        [Parameter(Mandatory)][string]$Path
    )
    Export-ReportCsv -Row $Row -Path $Path
}

<#
.SYNOPSIS
    Renders an inventory (or a subset of its rows) as a self-contained HTML dashboard.
#>
function ConvertTo-PimInventoryHtml {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Inventory,
        [AllowEmptyCollection()][object[]]$Row
    )
    $selected = if ($PSBoundParameters.ContainsKey('Row')) { $Row } else { $Inventory.Rows }
    ConvertTo-HtmlReport -Row @($selected) -Title $Inventory.Title -WarningDays $Inventory.ExpiryWarningDays `
        -Now $Inventory.GeneratedUtc -ScopeError @($Inventory.ScopeErrors)
}

<#
.SYNOPSIS
    Writes the HTML dashboard for an inventory to a file.
#>
function Export-PimInventoryHtml {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Inventory,
        [Parameter(Mandatory)][string]$Path,
        [AllowEmptyCollection()][object[]]$Row
    )
    $html = if ($PSBoundParameters.ContainsKey('Row')) {
        ConvertTo-PimInventoryHtml -Inventory $Inventory -Row $Row
    }
    else {
        ConvertTo-PimInventoryHtml -Inventory $Inventory
    }
    [System.IO.File]::WriteAllText($Path, $html, [System.Text.UTF8Encoding]::new($false))
}

#endregion

Export-ModuleMember -Function @(
    'Get-PimInventory'
    'Get-PimManagementGroup'
    'Export-PimInventoryCsv'
    'ConvertTo-PimInventoryHtml'
    'Export-PimInventoryHtml'
)
