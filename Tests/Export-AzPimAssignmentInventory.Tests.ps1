#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.3.0' }

BeforeAll {
    $script:RepoRoot        = Split-Path -Parent $PSScriptRoot
    $script:InventoryScript = Join-Path $script:RepoRoot 'Export-AzPimAssignmentInventory.ps1'
    $script:SampleData      = Join-Path $script:RepoRoot 'sample-data.json'

    # Load every function from the script without running its main body, so units
    # can be tested in isolation. Parse errors fail the whole suite immediately.
    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile(
        $script:InventoryScript, [ref]$tokens, [ref]$parseErrors
    )
    if ($parseErrors.Count -gt 0) { throw ($parseErrors -join [Environment]::NewLine) }

    $functions = $ast.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst]
        }, $false)
    foreach ($f in $functions) { . ([scriptblock]::Create($f.Extent.Text)) }

    # Script-scope state the functions expect.
    $script:ArmApiVersion = '2020-10-01'
    $script:MgApiVersion  = '2020-05-01'
    $script:ReportColumns = @(
        'PrincipalName', 'PrincipalType', 'PrincipalEmail', 'PrincipalId',
        'RoleName', 'RoleType', 'PrivilegedRole', 'AssignmentKind', 'MemberType',
        'SubscriptionName', 'ScopeName', 'ScopeType', 'Scope', 'ResourceUri',
        'StartDateUtc', 'ExpiryDateUtc', 'DaysRemaining', 'Status', 'Condition', 'InstanceId'
    )

    # Stub so Az.Accounts is not needed on the test machine; every test mocks it.
    function Invoke-AzRestMethod {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '', Justification = 'Signature stub for Mock.')]
        param($Path, $Method)
        throw 'Azure requests must be mocked.'
    }

    function Get-TestInstance {
        param(
            [string]$Id = '/subscriptions/s1/providers/Microsoft.Authorization/RoleEligibilityScheduleInstances/1',
            [string]$Kind = 'Eligible',
            [string]$Scope = '/subscriptions/s1',
            [string]$ScopeName = 'Sub One',
            [string]$RoleName = 'Reader',
            [string]$PrincipalName = 'SG-Test',
            $Start = '2026-01-01T00:00:00Z',
            $End = $null,
            [string]$AssignmentType = ''
        )
        $properties = [ordered]@{
            scope              = $Scope
            principalId        = 'p1'
            principalType      = 'Group'
            memberType         = 'Direct'
            startDateTime      = $Start
            endDateTime        = $End
            expandedProperties = [pscustomobject]@{
                scope          = [pscustomobject]@{ id = $Scope; displayName = $ScopeName }
                roleDefinition = [pscustomobject]@{ displayName = $RoleName; type = 'BuiltInRole' }
                principal      = [pscustomobject]@{ displayName = $PrincipalName; email = ''; type = 'Group' }
            }
        }
        if ($AssignmentType) { $properties.assignmentType = $AssignmentType }
        [pscustomobject]@{ _kind = $Kind; id = $Id; properties = [pscustomobject]$properties }
    }

    $script:Now = [datetime]::SpecifyKind([datetime]'2026-09-11T00:00:00', 'Utc')
}

Describe 'Get-TargetScope' {
    It 'keeps the root as a collection when no descendants are returned' {
        Mock Invoke-ArmList { return , @() }

        $scopes = Get-TargetScope -ManagementGroupId 'mg-test'

        $scopes.Count | Should -Be 1
        $scopes[0] | Should -Be '/providers/Microsoft.Management/managementGroups/mg-test'
        Should -Invoke Invoke-ArmList -Times 1 -Exactly -ParameterFilter {
            $Path -eq '/providers/Microsoft.Management/managementGroups/mg-test/descendants?api-version=2020-05-01'
        }
    }

    It 'preserves descendant traversal, removes duplicates and captures subscription names' {
        Mock Invoke-ArmList {
            return , @(
                [pscustomobject]@{ type = 'Microsoft.Management/managementGroups'; name = 'mg-child' }
                [pscustomobject]@{ type = 'Microsoft.Management/managementGroups/subscriptions'; name = 'sub-3'; properties = [pscustomobject]@{ displayName = 'Prod' } }
                [pscustomobject]@{ type = 'Microsoft.Management/managementGroups/subscriptions'; name = 'sub-3'; properties = [pscustomobject]@{ displayName = 'Prod' } }
            )
        }
        $names = @{}

        $scopes = Get-TargetScope -ManagementGroupId 'mg-test' -SubscriptionNames $names

        $scopes | Should -Be @(
            '/providers/Microsoft.Management/managementGroups/mg-test'
            '/providers/Microsoft.Management/managementGroups/mg-child'
            '/subscriptions/sub-3'
        )
        $names['sub-3'] | Should -Be 'Prod'
    }
}

Describe 'Invoke-ArmList' {
    BeforeEach { Mock Start-Sleep { } }

    It 'follows nextLink paging and strips the host' {
        Mock Invoke-AzRestMethod {
            [pscustomobject]@{ StatusCode = 200; Content = '{"value":[{"id":"a"}],"nextLink":"https://management.azure.com/next?page=2"}' }
        } -ParameterFilter { $Path -eq '/first' }
        Mock Invoke-AzRestMethod {
            [pscustomobject]@{ StatusCode = 200; Content = '{"value":[{"id":"b"}]}' }
        } -ParameterFilter { $Path -eq '/next?page=2' }

        $result = Invoke-ArmList -Path '/first'

        $result.id | Should -Be @('a', 'b')
    }

    It 'retries on 429 and honours Retry-After' {
        $script:calls = 0
        Mock Invoke-AzRestMethod {
            $script:calls++
            if ($script:calls -eq 1) {
                return [pscustomobject]@{ StatusCode = 429; Content = ''; Headers = @{ 'Retry-After' = @('7') } }
            }
            [pscustomobject]@{ StatusCode = 200; Content = '{"value":[{"id":"x"}]}' }
        }

        $result = Invoke-ArmList -Path '/throttled'

        $result.id | Should -Be 'x'
        Should -Invoke Start-Sleep -Times 1 -Exactly -ParameterFilter { $Seconds -eq 7 }
    }

    It 'throws on a non-transient failure without retrying' {
        Mock Invoke-AzRestMethod { [pscustomobject]@{ StatusCode = 403; Content = 'Forbidden' } }

        { Invoke-ArmList -Path '/denied' } | Should -Throw '*403*'
        Should -Invoke Invoke-AzRestMethod -Times 1 -Exactly
    }

    It 'gives up after MaxRetry transient failures' {
        Mock Invoke-AzRestMethod { [pscustomobject]@{ StatusCode = 503; Content = '' } }

        { Invoke-ArmList -Path '/down' -MaxRetry 3 } | Should -Throw '*503 after 3 attempts*'
    }
}

Describe 'Get-RetryDelay' {
    It 'uses delta-seconds when given' {
        Get-RetryDelay -Response ([pscustomobject]@{ Headers = @{ 'Retry-After' = '12' } }) | Should -Be 12
    }

    It 'accepts an HTTP-date without throwing' {
        $date = [DateTimeOffset]::UtcNow.AddSeconds(20).ToString('r')
        $delay = Get-RetryDelay -Response ([pscustomobject]@{ Headers = @{ 'Retry-After' = $date } })
        $delay | Should -BeGreaterOrEqual 1
        $delay | Should -BeLessOrEqual 21
    }

    It 'falls back to exponential backoff for garbage or missing headers' {
        Get-RetryDelay -Response ([pscustomobject]@{ Headers = @{ 'Retry-After' = 'soon' } }) -Attempt 3 | Should -Be 8
        Get-RetryDelay -Response $null -Attempt 2 | Should -Be 4
    }
}

Describe 'Get-PimInstance' {
    It 'queries schedule instances by default and tags the kind' {
        Mock Invoke-ArmList { return , @([pscustomobject]@{ id = 'i1' }) }

        $items = Get-PimInstance -Scope '/subscriptions/s1' -Kind 'Eligible'

        $items[0]._kind | Should -Be 'Eligible'
        Should -Invoke Invoke-ArmList -ParameterFilter { $Path -like '*/roleEligibilityScheduleInstances?api-version=2020-10-01' }
    }

    It 'queries schedules when future assignments are requested' {
        Mock Invoke-ArmList { return , @() }

        $null = Get-PimInstance -Scope '/subscriptions/s1' -Kind 'Active' -IncludeFuture

        Should -Invoke Invoke-ArmList -ParameterFilter { $Path -like '*/roleAssignmentSchedules?api-version=*' }
    }

    It 'records a scope failure and continues' {
        Mock Invoke-ArmList { throw 'boom' }
        $log = [System.Collections.Generic.List[object]]::new()

        $items = Get-PimInstance -Scope '/subscriptions/s1' -Kind 'Active' -ErrorLog $log -WarningAction SilentlyContinue

        @($items).Count | Should -Be 0
        $log.Count | Should -Be 1
        $log[0].Message | Should -Be 'boom'
    }
}

Describe 'Get-UniqueInstance' {
    It 'de-duplicates on id case-insensitively and keeps id-less records' {
        $result = Get-UniqueInstance -Instance @(
            [pscustomobject]@{ id = '/A/1' }
            [pscustomobject]@{ id = '/a/1' }
            [pscustomobject]@{ id = '/a/2' }
            [pscustomobject]@{ name = 'no-id' }
        )
        $result.Count | Should -Be 3
    }
}

Describe 'Get-ScopeType' {
    It 'classifies <Scope> as <Expected>' -ForEach @(
        @{ Scope = '/providers/Microsoft.Management/managementGroups/mg1'; Expected = 'Management group' }
        @{ Scope = '/subscriptions/s1'; Expected = 'Subscription' }
        @{ Scope = '/subscriptions/s1/resourceGroups/rg1'; Expected = 'Resource group' }
        @{ Scope = '/subscriptions/s1/resourceGroups/rg1/providers/Microsoft.Web/sites/app'; Expected = 'Resource' }
        @{ Scope = '/'; Expected = 'Tenant root' }
        @{ Scope = ''; Expected = 'Unknown' }
    ) {
        Get-ScopeType $Scope | Should -Be $Expected
    }
}

Describe 'ConvertTo-UtcDateTime' {
    It 'parses an ISO 8601 Z string as UTC' {
        $d = ConvertTo-UtcDateTime '2026-09-25T00:00:00Z'
        $d.Kind | Should -Be 'Utc'
        $d.ToString('yyyy-MM-dd HH:mm') | Should -Be '2026-09-25 00:00'
    }

    It 'honours an explicit offset' {
        (ConvertTo-UtcDateTime '2026-09-25T10:30:00+10:30').ToString('yyyy-MM-dd HH:mm') | Should -Be '2026-09-25 00:00'
    }

    It 'converts a local DateTime (PowerShell 7 ConvertFrom-Json output) to UTC' {
        $utc = [datetime]::SpecifyKind([datetime]'2026-09-25T00:00:00', 'Utc')
        (ConvertTo-UtcDateTime $utc.ToLocalTime()) | Should -Be $utc
    }

    It 'returns null for empty or unparseable input' {
        ConvertTo-UtcDateTime $null | Should -BeNullOrEmpty
        ConvertTo-UtcDateTime '' | Should -BeNullOrEmpty
        ConvertTo-UtcDateTime 'not a date' | Should -BeNullOrEmpty
    }
}

Describe 'ConvertTo-ReportRow' {
    It 'computes status <Expected> for end <End> and start <Start>' -ForEach @(
        @{ Start = '2026-01-01T00:00:00Z'; End = $null; Expected = 'Permanent' }
        @{ Start = '2026-01-01T00:00:00Z'; End = '2026-09-10T00:00:00Z'; Expected = 'Expired' }
        @{ Start = '2026-01-01T00:00:00Z'; End = '2026-09-11T06:00:00Z'; Expected = 'Expiring' }
        @{ Start = '2026-01-01T00:00:00Z'; End = '2026-10-11T00:00:00Z'; Expected = 'Expiring' }
        @{ Start = '2026-01-01T00:00:00Z'; End = '2026-10-12T00:00:00Z'; Expected = 'Active' }
        @{ Start = '2026-12-01T00:00:00Z'; End = '2027-12-01T00:00:00Z'; Expected = 'Pending' }
        @{ Start = '2026-12-01T00:00:00Z'; End = $null; Expected = 'Pending' }
        @{ Start = '2026-01-01T00:00:00Z'; End = 'garbage'; Expected = 'Unknown' }
    ) {
        $row = ConvertTo-ReportRow -Instance @(Get-TestInstance -Start $Start -End $End) -WarningDays 30 -Now $script:Now -WarningAction SilentlyContinue
        $row[0].Status | Should -Be $Expected
    }

    It 'reports 0 days remaining (not blank) for an assignment expiring later today' {
        $row = ConvertTo-ReportRow -Instance @(Get-TestInstance -End '2026-09-11T06:00:00Z') -WarningDays 30 -Now $script:Now
        $row[0].DaysRemaining | Should -Be 0
    }

    It 'classifies eligible, active and JIT assignments' {
        $rows = ConvertTo-ReportRow -WarningDays 30 -Now $script:Now -Instance @(
            Get-TestInstance -Id 'e' -Kind 'Eligible'
            Get-TestInstance -Id 'a' -Kind 'Active' -AssignmentType 'Assigned'
            Get-TestInstance -Id 'j' -Kind 'Active' -AssignmentType 'Activated'
        )
        $rows.AssignmentKind | Should -Be @('Eligible', 'Active', 'Activation (JIT)')
    }

    It 'flags privileged roles case-insensitively' {
        $rows = ConvertTo-ReportRow -WarningDays 30 -Now $script:Now -PrivilegedRoleName @('owner') -Instance @(
            Get-TestInstance -Id '1' -RoleName 'Owner'
            Get-TestInstance -Id '2' -RoleName 'Reader'
        )
        $rows.PrivilegedRole | Should -Be @('Yes', 'No')
    }

    It 'resolves subscription names for child scopes from a subscription-level row' {
        $rows = ConvertTo-ReportRow -WarningDays 30 -Now $script:Now -Instance @(
            Get-TestInstance -Id '1' -Scope '/subscriptions/s9/resourceGroups/rg' -ScopeName 'rg'
            Get-TestInstance -Id '2' -Scope '/subscriptions/s9' -ScopeName 'Finance'
        )
        $rows.SubscriptionName | Should -Be @('Finance', 'Finance')
    }

    It 'prefers hierarchy metadata for subscription names' {
        $rows = ConvertTo-ReportRow -WarningDays 30 -Now $script:Now -SubscriptionNames @{ s1 = 'From MG' } -Instance @(
            Get-TestInstance -Scope '/subscriptions/s1' -ScopeName 'From assignment'
        )
        $rows[0].SubscriptionName | Should -Be 'From MG'
    }

    It 'tolerates a record with no expandedProperties' {
        $bare = [pscustomobject]@{ _kind = 'Active'; id = 'x'; properties = [pscustomobject]@{ scope = '/providers/Microsoft.Management/managementGroups/mg' } }
        $row = ConvertTo-ReportRow -Instance @($bare) -WarningDays 30 -Now $script:Now
        $row[0].PrincipalName | Should -Be '(unresolved)'
        $row[0].RoleName | Should -Be '(unknown role)'
        $row[0].Status | Should -Be 'Permanent'
    }

    It 'emits exactly the documented columns in order' {
        $row = ConvertTo-ReportRow -Instance @(Get-TestInstance) -WarningDays 30 -Now $script:Now
        $row[0].PSObject.Properties.Name | Should -Be $script:ReportColumns
    }
}

Describe 'Select-InHierarchyRow' {
    BeforeAll {
        $script:HierarchyScopes = @(
            '/providers/Microsoft.Management/managementGroups/mg-prod'
            '/subscriptions/s1'
        )
    }

    It 'keeps <Scope> = <Expected>' -ForEach @(
        @{ Scope = '/providers/Microsoft.Management/managementGroups/mg-prod'; Expected = $true }
        @{ Scope = '/providers/microsoft.management/managementgroups/MG-PROD'; Expected = $true }
        @{ Scope = '/subscriptions/s1'; Expected = $true }
        @{ Scope = '/subscriptions/s1/resourceGroups/rg/providers/Microsoft.Web/sites/x'; Expected = $true }
        @{ Scope = '/providers/Microsoft.Management/managementGroups/mg-prod-legacy'; Expected = $false }
        @{ Scope = '/providers/Microsoft.Management/managementGroups/tenant-root'; Expected = $false }
        @{ Scope = '/subscriptions/s10'; Expected = $false }
        @{ Scope = '/'; Expected = $false }
    ) {
        $kept = Select-InHierarchyRow -Row @([pscustomobject]@{ Scope = $Scope }) -Scope $script:HierarchyScopes
        ($kept.Count -eq 1) | Should -Be $Expected
    }
}

Describe 'Protect-CsvValue' {
    It 'prefixes <Value>' -ForEach @(
        @{ Value = '=HYPERLINK("http://evil")'; Expected = "'=HYPERLINK(""http://evil"")" }
        @{ Value = '+cmd'; Expected = "'+cmd" }
        @{ Value = '-2+3'; Expected = "'-2+3" }
        @{ Value = '@SUM(A1)'; Expected = "'@SUM(A1)" }
        @{ Value = 'SG-Normal'; Expected = 'SG-Normal' }
        @{ Value = ''; Expected = '' }
    ) {
        Protect-CsvValue $Value | Should -Be $Expected
    }
}

Describe 'Export-ReportCsv' {
    It 'writes UTF-8 with BOM, neutralises formulas, keeps negative day counts numeric' {
        $path = Join-Path $TestDrive 'out.csv'
        $row = ConvertTo-ReportRow -Instance @(Get-TestInstance -PrincipalName '=evil()' -End '2026-09-01T00:00:00Z') -WarningDays 30 -Now $script:Now

        Export-ReportCsv -Row $row -Path $path

        $bytes = [System.IO.File]::ReadAllBytes($path)
        $bytes[0..2] | Should -Be @(0xEF, 0xBB, 0xBF)
        $csv = Import-Csv -LiteralPath $path
        $csv[0].PrincipalName | Should -Be "'=evil()"
        $csv[0].DaysRemaining | Should -Be '-10'
    }

    It 'writes a header row even when there are no assignments' {
        $path = Join-Path $TestDrive 'empty.csv'
        Export-ReportCsv -Row @() -Path $path
        (Get-Content -LiteralPath $path -First 1) | Should -BeLike '"PrincipalName","PrincipalType"*'
    }
}

Describe 'ConvertTo-HtmlReport' {
    It 'HTML-encodes untrusted display names' {
        $row = ConvertTo-ReportRow -Instance @(Get-TestInstance -PrincipalName '<script>alert(1)</script>') -WarningDays 30 -Now $script:Now
        $html = ConvertTo-HtmlReport -Row $row -Title 'T' -WarningDays 30 -Now $script:Now
        $html | Should -Not -Match '<script>alert'
        $html | Should -Match '&lt;script&gt;alert\(1\)&lt;/script&gt;'
    }

    It 'shows a banner when scopes could not be read' {
        $html = ConvertTo-HtmlReport -Row @() -Title 'T' -WarningDays 30 -ScopeError @(
            [pscustomobject]@{ Scope = '/subscriptions/s1'; Kind = 'Active'; Message = '403' }
        )
        $html | Should -Match '1 scope\(s\) could not be read'
    }

    It 'counts standing privileged access' {
        $rows = ConvertTo-ReportRow -WarningDays 30 -Now $script:Now -PrivilegedRoleName @('Owner') -Instance @(
            Get-TestInstance -Id '1' -Kind 'Active' -RoleName 'Owner' -AssignmentType 'Assigned'
            Get-TestInstance -Id '2' -Kind 'Eligible' -RoleName 'Owner'
        )
        $html = ConvertTo-HtmlReport -Row $rows -Title 'T' -WarningDays 30 -Now $script:Now
        $html | Should -Match '<div class="n">1</div><div class="l">Standing privileged access'
    }
}

Describe 'End-to-end offline run' {
    BeforeAll {
        $script:OutDir = Join-Path $TestDrive 'e2e'
        $script:Rows = & $script:InventoryScript -OfflineSampleData $script:SampleData -OutputPath $script:OutDir `
            -ReferenceDateUtc '2026-09-11 01:36' -PassThru 6>$null 3>$null
    }

    It 'de-duplicates and excludes JIT activations by default' {
        $script:Rows.Count | Should -Be 10
        $script:Rows.AssignmentKind | Should -Not -Contain 'Activation (JIT)'
    }

    It 'produces the expected status mix for the sample data' {
        ($script:Rows | Group-Object Status -AsHashTable -AsString).Keys | Sort-Object | Should -Be @('Active', 'Expired', 'Expiring', 'Permanent')
        @($script:Rows | Where-Object Status -eq 'Expiring').Count | Should -Be 5
    }

    It 'writes CSV and HTML files named after the reference date' {
        Test-Path (Join-Path $script:OutDir 'PIM-Assignment-Inventory-20260911-0136.csv') | Should -BeTrue
        Test-Path (Join-Path $script:OutDir 'PIM-Assignment-Inventory-20260911-0136.html') | Should -BeTrue
    }

    It 'includes JIT activations when asked' {
        $rows = & $script:InventoryScript -OfflineSampleData $script:SampleData -OutputPath $script:OutDir `
            -ReferenceDateUtc '2026-09-11 01:36' -PassThru -IncludeActivations -SkipHtml 6>$null 3>$null
        $rows.Count | Should -Be 11
    }
}
