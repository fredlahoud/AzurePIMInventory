<#
.SYNOPSIS
    Inventories every Azure PIM role assignment (eligible and active) beneath a
    management group, and exports the result to CSV and a self-contained HTML dashboard.

.DESCRIPTION
    Answers the question "show me every PIM assignment across all of my subscriptions
    and the expiry date", which the Azure portal cannot answer in a single view.

    Design notes that matter:

    * ARM only. This script calls the Azure Resource Manager Authorization provider.
      It makes NO Microsoft Graph calls, so it needs no Entra directory permissions
      and no tenant-level admin consent. Friendly names for the principal, role and
      scope come from the ARM 'expandedProperties' payload. Subscription names
      come from management-group descendant metadata.

    * Coverage is explicit, not assumed. The ARM list-for-scope operation returns
      assignments at, above and below the requested scope, but that downward
      traversal from a management group root is not documented by Microsoft. Rather
      than rely on it, this script enumerates the hierarchy itself (root management
      group + all descendant management groups + all subscriptions), queries each
      scope, and de-duplicates on the instance id. Coverage is therefore correct
      whether or not the traversal behaves as expected.

    * Both assignment surfaces are collected:
        - roleEligibilityScheduleInstances  -> eligible assignments
        - roleAssignmentScheduleInstances   -> active assignments, which includes
          permanent RBAC granted outside PIM entirely.

    * Just-in-time activations are excluded by default. An activated eligible
      assignment appears as an active instance with assignmentType 'Activated' and
      would otherwise pollute the inventory with rows that vanish within hours.
      Use -IncludeActivations to keep them.

    * Schedule INSTANCES only cover assignments that are active right now.
      Assignments scheduled to begin in the future live in the *Schedules resources
      instead, and are invisible by default. Use -IncludeFutureAssignments to query
      those instead, per Microsoft: "schedule instances only include assignments
      that are active at the current time, schedules also include assignments that
      will become active in the future".

.PARAMETER RootManagementGroupId
    Management group ID (the name, not the display name) to report from.

.PARAMETER OutputPath
    Directory for the CSV and HTML output. Defaults to the current directory.

.PARAMETER ExpiryWarningDays
    Assignments expiring within this many days are flagged 'Expiring'. Default 30.

.PARAMETER IncludeActivations
    Include transient just-in-time activations in the output.

.PARAMETER IncludeAncestorScopes
    Include assignments inherited from scopes ABOVE the root management group.
    Off by default, because in a shared tenant those are other agencies' assignments.

.PARAMETER IncludeFutureAssignments
    Query the *Schedules resources rather than *ScheduleInstances, so that
    assignments scheduled to start in the future are included as well as current
    ones. Off by default, which reports the estate as it stands today.

.PARAMETER SkipHtml
    Produce the CSV only.

.PARAMETER OfflineSampleData
    Path to a JSON file of sample instances, used to preview the output format
    without connecting to Azure. Intended for validation and demonstration.

.EXAMPLE
    .\Export-AzPimAssignmentInventory.ps1 -RootManagementGroupId 'mg-pirsa-root'

.EXAMPLE
    .\Export-AzPimAssignmentInventory.ps1 -RootManagementGroupId 'mg-pirsa-root' `
        -ExpiryWarningDays 60 -OutputPath 'C:\Reports'

.NOTES
    Requires : Az.Accounts (Connect-AzAccount first). Az.Resources is not required.
    Rights   : Reader at the root management group is normally sufficient, because
               listing needs Microsoft.Authorization/*/read. Microsoft documents
               Owner or User Access Administrator for the PIM ARM API generally, so
               if calls return 403, grant User Access Administrator instead.
    Licence  : PIM requires Microsoft Entra ID P2 or Microsoft Entra ID Governance.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory, ParameterSetName = 'Live')]
    [ValidateNotNullOrEmpty()]
    [string]$RootManagementGroupId,

    [Parameter(Mandatory, ParameterSetName = 'Offline')]
    [ValidateNotNullOrEmpty()]
    [string]$OfflineSampleData,

    [string]$OutputPath = (Get-Location).Path,

    [ValidateRange(1, 365)]
    [int]$ExpiryWarningDays = 30,

    [switch]$IncludeActivations,

    [switch]$IncludeAncestorScopes,

    [switch]$IncludeFutureAssignments,

    [switch]$SkipHtml
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ArmApiVersion = '2020-10-01'
$script:MgApiVersion  = '2020-05-01'
$script:ScopeErrors   = [System.Collections.Generic.List[object]]::new()

#region Helpers ---------------------------------------------------------------

function Write-Step {
    param([string]$Message)
    Write-Host "  $Message" -ForegroundColor Cyan
}

function Assert-Prerequisites {
    if (-not (Get-Module -ListAvailable -Name Az.Accounts)) {
        throw "The Az.Accounts module is required. Install it with: Install-Module Az.Accounts -Scope CurrentUser"
    }
    Import-Module Az.Accounts -ErrorAction Stop

    $context = Get-AzContext -ErrorAction SilentlyContinue
    if (-not $context) {
        throw "No Azure context found. Run Connect-AzAccount first."
    }
    Write-Step "Signed in as $($context.Account.Id) on tenant $($context.Tenant.Id)"
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
                Start-Sleep -Seconds ([math]::Min(30, [math]::Pow(2, $attempt)))
                continue
            }

            # 429 and 5xx are transient. Honour Retry-After when present.
            if ($response.StatusCode -eq 429 -or $response.StatusCode -ge 500) {
                if ($attempt -ge $MaxRetry) {
                    throw "ARM returned $($response.StatusCode) after $attempt attempts for $next"
                }
                $wait = 5
                if ($response.PSObject.Properties.Name -contains 'Headers' -and $response.Headers) {
                    $retryAfter = $response.Headers | Where-Object { $_.Key -eq 'Retry-After' } | Select-Object -First 1
                    if ($retryAfter -and $retryAfter.Value) { $wait = [int]($retryAfter.Value | Select-Object -First 1) }
                }
                Start-Sleep -Seconds $wait
                continue
            }

            break
        }

        if ($response.StatusCode -ne 200) {
            throw "ARM returned $($response.StatusCode) for $next : $($response.Content)"
        }

        $payload = $response.Content | ConvertFrom-Json
        if ($payload.PSObject.Properties.Name -contains 'value' -and $payload.value) {
            foreach ($item in $payload.value) { $results.Add($item) }
        }

        $next = $null
        if ($payload.PSObject.Properties.Name -contains 'nextLink' -and $payload.nextLink) {
            # nextLink is an absolute URL; Invoke-AzRestMethod -Path wants a relative path.
            $next = $payload.nextLink -replace '^https://[^/]+', ''
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
            $displayName = Get-SafeProperty (Get-SafeProperty $d 'properties') 'displayName' ''
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
        [Parameter(Mandatory)][ValidateSet('Eligible', 'Active')][string]$Kind
    )

    # Schedule instances are assignments active now. Schedules also include
    # assignments scheduled to start in the future.
    if ($IncludeFutureAssignments) {
        $resource = if ($Kind -eq 'Eligible') { 'roleEligibilitySchedules' } else { 'roleAssignmentSchedules' }
    }
    else {
        $resource = if ($Kind -eq 'Eligible') { 'roleEligibilityScheduleInstances' } else { 'roleAssignmentScheduleInstances' }
    }
    $path = "$Scope/providers/Microsoft.Authorization/$resource" + "?api-version=$script:ArmApiVersion"

    try {
        $items = Invoke-ArmList -Path $path
    }
    catch {
        # A single unreadable scope must not abort the whole inventory. Record and continue.
        $script:ScopeErrors.Add([pscustomobject]@{
                Scope   = $Scope
                Kind    = $Kind
                Message = $_.Exception.Message
            })
        Write-Warning "Could not read $Kind assignments at $Scope : $($_.Exception.Message)"
        return , @()
    }

    foreach ($i in $items) {
        Add-Member -InputObject $i -NotePropertyName '_kind' -NotePropertyValue $Kind -Force
    }
    return , $items
}

#endregion

#region Transform -------------------------------------------------------------

function Get-ScopeType {
    param([string]$Scope)

    if ([string]::IsNullOrWhiteSpace($Scope)) { return 'Unknown' }
    switch -Regex ($Scope) {
        '/providers/Microsoft\.Management/managementGroups/[^/]+$' { return 'Management group' }
        '^/subscriptions/[^/]+$'                                   { return 'Subscription' }
        '^/subscriptions/[^/]+/resourceGroups/[^/]+$'              { return 'Resource group' }
        '^/subscriptions/[^/]+/resourceGroups/[^/]+/providers/.+'  { return 'Resource' }
        default                                                    { return 'Other' }
    }
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
    Flattens raw ARM instances into report rows and computes expiry state.
    Kept separate from the ARM calls so it can be exercised with sample data.
#>
function ConvertTo-ReportRow {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Instance,
        [Parameter(Mandatory)][int]$WarningDays,
        [Parameter(Mandatory)][datetime]$Now,
        [hashtable]$SubscriptionNames = @{}
    )

    $rows = [System.Collections.Generic.List[object]]::new()
    $names = $SubscriptionNames.Clone()

    # Resolve offline subscription names before processing any child-scope rows.
    foreach ($i in $Instance) {
        $p = Get-SafeProperty $i 'properties'
        $scopeInfo = Get-SafeProperty (Get-SafeProperty $p 'expandedProperties') 'scope'
        $scope = Get-SafeProperty $p 'scope' (Get-SafeProperty $scopeInfo 'id' '')
        if ($scope -match '^/subscriptions/([^/]+)$') {
            $subscriptionId = $Matches[1]
            $displayName = Get-SafeProperty $scopeInfo 'displayName' ''
            if (-not [string]::IsNullOrWhiteSpace($displayName) -and
                (-not $names.ContainsKey($subscriptionId) -or
                 [string]::IsNullOrWhiteSpace($names[$subscriptionId]))) {
                $names[$subscriptionId] = $displayName
            }
        }
    }

    foreach ($i in $Instance) {
        $p        = Get-SafeProperty $i 'properties'
        $expanded = Get-SafeProperty $p 'expandedProperties'

        $principal      = Get-SafeProperty $expanded 'principal'
        $roleDefinition = Get-SafeProperty $expanded 'roleDefinition'
        $scopeInfo      = Get-SafeProperty $expanded 'scope'

        $scope = Get-SafeProperty $p 'scope' (Get-SafeProperty $scopeInfo 'id' '')
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

        $endRaw   = Get-SafeProperty $p 'endDateTime'
        $startRaw = Get-SafeProperty $p 'startDateTime'

        $end = $null
        if ($endRaw) { $end = [datetime]::Parse($endRaw, [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AdjustToUniversal) }
        $start = $null
        if ($startRaw) { $start = [datetime]::Parse($startRaw, [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AdjustToUniversal) }

        $daysRemaining = $null
        if ($end) { $daysRemaining = [math]::Floor(($end - $Now).TotalDays) }

        $status = if (-not $end) { 'Permanent' }
        elseif ($daysRemaining -lt 0) { 'Expired' }
        elseif ($daysRemaining -le $WarningDays) { 'Expiring' }
        else { 'Active' }

        $kind           = Get-SafeProperty $i '_kind' 'Unknown'
        $assignmentType = Get-SafeProperty $p 'assignmentType' ''

        # Eligible instances have no assignmentType. Active ones are either standing
        # access ('Assigned') or a live just-in-time activation ('Activated').
        $category = if ($kind -eq 'Eligible') { 'Eligible' }
        elseif ($assignmentType -eq 'Activated') { 'Activation (JIT)' }
        else { 'Active' }

        $rows.Add([pscustomobject]@{
                PrincipalName   = Get-SafeProperty $principal 'displayName' '(unresolved)'
                PrincipalType   = Get-SafeProperty $p 'principalType' (Get-SafeProperty $principal 'type' '')
                PrincipalEmail  = Get-SafeProperty $principal 'email' ''
                PrincipalId     = Get-SafeProperty $p 'principalId' ''
                RoleName        = Get-SafeProperty $roleDefinition 'displayName' '(unknown role)'
                RoleType        = Get-SafeProperty $roleDefinition 'type' ''
                AssignmentKind  = $category
                MemberType      = Get-SafeProperty $p 'memberType' ''
                SubscriptionName = $subscriptionName
                ScopeName       = Get-SafeProperty $scopeInfo 'displayName' ''
                ScopeType       = Get-ScopeType $scope
                Scope           = $scope
                ResourceUri     = $scope
                StartDateUtc    = if ($start) { $start.ToString('yyyy-MM-dd HH:mm') } else { '' }
                ExpiryDateUtc   = if ($end) { $end.ToString('yyyy-MM-dd HH:mm') } else { '' }
                DaysRemaining   = if ($null -ne $daysRemaining) { $daysRemaining } else { '' }
                Status          = $status
                Condition       = if (Get-SafeProperty $p 'condition' '') { 'Yes' } else { 'No' }
                InstanceId      = Get-SafeProperty $i 'id' ''
            })
    }

    return , $rows.ToArray()
}

#endregion

#region Output ----------------------------------------------------------------

function ConvertTo-HtmlReport {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Row,
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][int]$WarningDays,
        [AllowEmptyCollection()][object[]]$ScopeError = @()
    )

    $total     = $Row.Count
    $expired   = @($Row | Where-Object { $_.Status -eq 'Expired' }).Count
    $expiring  = @($Row | Where-Object { $_.Status -eq 'Expiring' }).Count
    $permanent = @($Row | Where-Object { $_.Status -eq 'Permanent' }).Count
    $eligible  = @($Row | Where-Object { $_.AssignmentKind -eq 'Eligible' }).Count
    $active    = @($Row | Where-Object { $_.AssignmentKind -eq 'Active' }).Count

    $sb = [System.Text.StringBuilder]::new()

    # Sort so the rows that need action appear first.
    # Test for a blank with IsNullOrEmpty, not -eq '': PowerShell coerces the empty
    # string to 0, so "0 -eq ''" is true and an assignment expiring today would
    # otherwise sort to the bottom alongside the permanent ones.
    $order  = @{ 'Expired' = 0; 'Expiring' = 1; 'Active' = 2; 'Permanent' = 3 }
    $sorted = $Row | Sort-Object `
        @{ Expression = { $order[$_.Status] } }, `
        @{ Expression = { if ([string]::IsNullOrEmpty([string]$_.DaysRemaining)) { [int]::MaxValue } else { [int]$_.DaysRemaining } } }

    foreach ($r in $sorted) {
        $cls = switch ($r.Status) {
            'Expired'   { 'st-expired' }
            'Expiring'  { 'st-expiring' }
            'Permanent' { 'st-permanent' }
            default     { 'st-ok' }
        }
        [void]$sb.Append('<tr>')
        [void]$sb.Append("<td>$([System.Net.WebUtility]::HtmlEncode([string]$r.PrincipalName))</td>")
        [void]$sb.Append("<td>$([System.Net.WebUtility]::HtmlEncode([string]$r.PrincipalType))</td>")
        [void]$sb.Append("<td>$([System.Net.WebUtility]::HtmlEncode([string]$r.RoleName))</td>")
        [void]$sb.Append("<td>$([System.Net.WebUtility]::HtmlEncode([string]$r.AssignmentKind))</td>")
        [void]$sb.Append("<td>$([System.Net.WebUtility]::HtmlEncode([string]$r.SubscriptionName))</td>")
        [void]$sb.Append("<td>$([System.Net.WebUtility]::HtmlEncode([string]$r.ScopeType))</td>")
        [void]$sb.Append("<td title=""$([System.Net.WebUtility]::HtmlEncode([string]$r.Scope))"">$([System.Net.WebUtility]::HtmlEncode([string]$r.ScopeName))</td>")
        [void]$sb.Append("<td class=""resource-uri"">$([System.Net.WebUtility]::HtmlEncode([string]$r.ResourceUri))</td>")
        [void]$sb.Append("<td>$([System.Net.WebUtility]::HtmlEncode([string]$r.ExpiryDateUtc))</td>")
        [void]$sb.Append("<td class=""num"">$([System.Net.WebUtility]::HtmlEncode([string]$r.DaysRemaining))</td>")
        [void]$sb.Append("<td><span class=""pill $cls"">$([System.Net.WebUtility]::HtmlEncode([string]$r.Status))</span></td>")
        [void]$sb.Append('</tr>')
    }

    $errorBlock = ''
    if ($ScopeError.Count -gt 0) {
        $items = ($ScopeError | ForEach-Object {
                "<li>$([System.Net.WebUtility]::HtmlEncode([string]$_.Scope)) ($([System.Net.WebUtility]::HtmlEncode([string]$_.Kind))): $([System.Net.WebUtility]::HtmlEncode([string]$_.Message))</li>"
            }) -join ''
        $errorBlock = "<div class=""warn""><strong>$($ScopeError.Count) scope(s) could not be read.</strong> The inventory below is incomplete.<ul>$items</ul></div>"
    }

    $generated = (Get-Date).ToString('dd MMM yyyy HH:mm')
    $rowsHtml  = $sb.ToString()

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>$([System.Net.WebUtility]::HtmlEncode($Title))</title>
<style>
 :root { --ink:#1b1b1f; --muted:#5f6368; --line:#e3e3e6; --bg:#faf9f8; }
 * { box-sizing:border-box; }
 body { font-family:'Segoe UI',system-ui,-apple-system,sans-serif; margin:0; background:var(--bg); color:var(--ink); }
 header { background:#0f6cbd; color:#fff; padding:20px 28px; }
 header h1 { margin:0; font-size:19px; font-weight:600; }
 header p { margin:4px 0 0; font-size:13px; opacity:.9; }
 main { padding:22px 28px 60px; }
 .cards { display:flex; flex-wrap:wrap; gap:12px; margin-bottom:20px; }
 .card { background:#fff; border:1px solid var(--line); border-radius:8px; padding:14px 18px; min-width:132px; }
 .card .n { font-size:26px; font-weight:600; line-height:1.1; }
 .card .l { font-size:12px; color:var(--muted); margin-top:3px; }
 .card.bad .n { color:#a4262c; } .card.warnc .n { color:#a86200; } .card.perm .n { color:#7719aa; }
 .warn { background:#fff4ce; border:1px solid #e9c96a; border-radius:8px; padding:12px 16px; margin-bottom:18px; font-size:13px; }
 .warn ul { margin:8px 0 0 18px; }
 .toolbar { display:flex; gap:10px; margin-bottom:12px; flex-wrap:wrap; }
 input[type=search], select { padding:8px 11px; border:1px solid var(--line); border-radius:6px; font-size:13px; font-family:inherit; background:#fff; }
 input[type=search] { min-width:290px; }
 .table-wrap { overflow-x:auto; }
 table { width:100%; border-collapse:collapse; background:#fff; border:1px solid var(--line); border-radius:8px; overflow:hidden; font-size:13px; }
 th { background:#f3f2f1; text-align:left; padding:10px 12px; font-weight:600; cursor:pointer; white-space:nowrap; user-select:none; border-bottom:1px solid var(--line); }
 th:hover { background:#e9e8e7; }
 th::after { content:' \2195'; opacity:.35; font-size:11px; }
 td { padding:9px 12px; border-bottom:1px solid #f0f0f0; vertical-align:top; }
 td.num { text-align:right; font-variant-numeric:tabular-nums; }
 td.resource-uri { min-width:240px; overflow-wrap:anywhere; }
 tbody tr:hover { background:#f7f9fc; }
 .pill { display:inline-block; padding:2px 9px; border-radius:11px; font-size:11.5px; font-weight:600; white-space:nowrap; }
 .st-expired { background:#fde7e9; color:#a4262c; }
 .st-expiring { background:#fff4ce; color:#8a5700; }
 .st-ok { background:#dff6dd; color:#0b6a0b; }
 .st-permanent { background:#f3e8fb; color:#7719aa; }
 footer { margin-top:18px; font-size:12px; color:var(--muted); }
 .hidden { display:none; }
</style>
</head>
<body>
<header>
  <h1>$([System.Net.WebUtility]::HtmlEncode($Title))</h1>
  <p>Generated $generated &nbsp;|&nbsp; Expiry warning threshold: $WarningDays days &nbsp;|&nbsp; Source: Azure Resource Manager (no Microsoft Graph)</p>
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
</div>

<div class="toolbar">
  <input type="search" id="q" placeholder="Filter by principal, role, subscription, scope...">
  <select id="statusFilter">
    <option value="">All statuses</option>
    <option value="Expired">Expired</option>
    <option value="Expiring">Expiring</option>
    <option value="Active">Active</option>
    <option value="Permanent">Permanent</option>
  </select>
  <select id="kindFilter">
    <option value="">All assignment kinds</option>
    <option value="Eligible">Eligible</option>
    <option value="Active">Active</option>
    <option value="Activation (JIT)">Activation (JIT)</option>
  </select>
</div>

<div class="table-wrap">
<table id="t">
<thead><tr>
  <th>Principal</th><th>Type</th><th>Role</th><th>Assignment</th>
  <th>Subscription</th><th>Scope type</th><th>Scope</th><th>Resource URI</th>
  <th>Expiry (UTC)</th><th>Days left</th><th>Status</th>
</tr></thead>
<tbody>
$rowsHtml
</tbody>
</table>
</div>
<footer>Hover a scope name to see its full Azure resource path. Click any column heading to sort.</footer>
</main>
<script>
(function () {
  var table = document.getElementById('t');
  var body = table.tBodies[0];
  var rows = Array.prototype.slice.call(body.rows);
  var q = document.getElementById('q');
  var sf = document.getElementById('statusFilter');
  var kf = document.getElementById('kindFilter');

  function applyFilter() {
    var term = q.value.toLowerCase();
    var status = sf.value;
    var kind = kf.value;
    rows.forEach(function (r) {
      var text = r.textContent.toLowerCase();
      var okText = !term || text.indexOf(term) !== -1;
      var okStatus = !status || r.cells[10].textContent.trim() === status;
      var okKind = !kind || r.cells[3].textContent.trim() === kind;
      r.classList.toggle('hidden', !(okText && okStatus && okKind));
    });
  }
  q.addEventListener('input', applyFilter);
  sf.addEventListener('change', applyFilter);
  kf.addEventListener('change', applyFilter);

  var dir = {};
  Array.prototype.forEach.call(table.tHead.rows[0].cells, function (th, idx) {
    th.addEventListener('click', function () {
      dir[idx] = !dir[idx];
      var sign = dir[idx] ? 1 : -1;
      rows.sort(function (a, b) {
        var x = a.cells[idx].textContent.trim();
        var y = b.cells[idx].textContent.trim();
        var nx = parseFloat(x), ny = parseFloat(y);
        if (!isNaN(nx) && !isNaN(ny)) { return (nx - ny) * sign; }
        return x.localeCompare(y) * sign;
      });
      rows.forEach(function (r) { body.appendChild(r); });
    });
  });
})();
</script>
</body>
</html>
"@

    return $html
}

#endregion

#region Main ------------------------------------------------------------------

$now       = (Get-Date).ToUniversalTime()
$stamp     = $now.ToString('yyyyMMdd-HHmm')
$instances = [System.Collections.Generic.List[object]]::new()
$subscriptionNames = @{}

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

Write-Host ''
Write-Host 'Azure PIM assignment inventory' -ForegroundColor White

if ($PSCmdlet.ParameterSetName -eq 'Offline') {
    Write-Step "Offline preview using $OfflineSampleData"
    $raw = Get-Content -LiteralPath $OfflineSampleData -Raw | ConvertFrom-Json
    foreach ($i in $raw) { $instances.Add($i) }
    $reportTitle = 'Azure PIM assignment inventory (sample data)'
}
else {
    Assert-Prerequisites

    Write-Step "Enumerating hierarchy beneath '$RootManagementGroupId'"
    $scopes = Get-TargetScope -ManagementGroupId $RootManagementGroupId -SubscriptionNames $subscriptionNames
    $mgCount  = @($scopes | Where-Object { $_ -like '*managementGroups*' }).Count
    $subCount = @($scopes | Where-Object { $_ -like '/subscriptions/*' }).Count
    Write-Step "Found $mgCount management group(s) and $subCount subscription(s)"

    $n = 0
    foreach ($scope in $scopes) {
        $n++
        Write-Progress -Activity 'Collecting PIM assignments' -Status $scope -PercentComplete (($n / $scopes.Count) * 100)
        foreach ($kind in @('Eligible', 'Active')) {
            foreach ($item in (Get-PimInstance -Scope $scope -Kind $kind)) { $instances.Add($item) }
        }
    }
    Write-Progress -Activity 'Collecting PIM assignments' -Completed

    $reportTitle = "Azure PIM assignment inventory - $RootManagementGroupId"
}

# The same assignment is returned by several scope queries because ARM reports
# at, above and below each scope. De-duplicate on the instance id.
$unique = $instances | Group-Object { $_.id } | ForEach-Object { $_.Group[0] }
Write-Step "Collected $($instances.Count) record(s), $(@($unique).Count) unique after de-duplication"

$rows = ConvertTo-ReportRow -Instance @($unique) -WarningDays $ExpiryWarningDays -Now $now -SubscriptionNames $subscriptionNames

if (-not $IncludeActivations) {
    $before = $rows.Count
    $rows   = @($rows | Where-Object { $_.AssignmentKind -ne 'Activation (JIT)' })
    if ($before -ne $rows.Count) {
        Write-Step "Excluded $($before - $rows.Count) transient just-in-time activation(s). Use -IncludeActivations to keep them."
    }
}

if (-not $IncludeAncestorScopes -and $PSCmdlet.ParameterSetName -eq 'Live') {
    $rootScope = "/providers/Microsoft.Management/managementGroups/$RootManagementGroupId"
    $inScope   = @($scopes)
    $before    = $rows.Count
    $rows = @($rows | Where-Object {
            $s = $_.Scope
            $s -eq $rootScope -or ($inScope | Where-Object { $s -like "$_*" })
        })
    if ($before -ne $rows.Count) {
        Write-Step "Excluded $($before - $rows.Count) assignment(s) inherited from above the root. Use -IncludeAncestorScopes to keep them."
    }
}

$rows = @($rows)

if ($rows.Count -eq 0) {
    Write-Warning 'No PIM assignments were found. Verify the management group ID and that the signed-in identity can read role assignments at that scope.'
}

$csvPath = Join-Path $OutputPath "PIM-Assignment-Inventory-$stamp.csv"
$rows | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
Write-Step "CSV written to $csvPath"

if (-not $SkipHtml) {
    $htmlPath = Join-Path $OutputPath "PIM-Assignment-Inventory-$stamp.html"
    $html = ConvertTo-HtmlReport -Row $rows -Title $reportTitle -WarningDays $ExpiryWarningDays -ScopeError $script:ScopeErrors.ToArray()
    [System.IO.File]::WriteAllText($htmlPath, $html, [System.Text.UTF8Encoding]::new($false))
    Write-Step "HTML dashboard written to $htmlPath"
}

Write-Host ''
Write-Host 'Summary' -ForegroundColor White
$rows | Group-Object Status | Sort-Object Name | ForEach-Object {
    Write-Host ("  {0,-10} {1}" -f $_.Name, $_.Count)
}

if ($script:ScopeErrors.Count -gt 0) {
    Write-Host ''
    Write-Warning "$($script:ScopeErrors.Count) scope(s) could not be read. The inventory is incomplete - see the HTML report for detail."
}

Write-Host ''

#endregion
