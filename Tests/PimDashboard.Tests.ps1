#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.3.0' }

BeforeDiscovery {
    $script:IsWindowsHost = ($env:OS -eq 'Windows_NT')
}

BeforeAll {
    $script:RepoRoot   = Split-Path -Parent $PSScriptRoot
    $script:AppScript  = Join-Path $script:RepoRoot 'App/PimDashboard.ps1'
    $script:XamlFile   = Join-Path $script:RepoRoot 'App/MainWindow.xaml'
    $script:SampleData = Join-Path $script:RepoRoot 'sample-data.json'

    # Dot-sourcing loads the functions; the app only starts when executed directly.
    . $script:AppScript
    $script:LogPath = Join-Path $TestDrive 'app.log'
    $script:DataDir = $TestDrive

    Import-Module (Join-Path $script:RepoRoot 'src/AzPimInventory/AzPimInventory.psd1') -Force
    $script:SampleInventory = Get-PimInventory -OfflineSampleData $script:SampleData -ReferenceDateUtc '2026-09-11 01:36' 6>$null 3>$null
    $script:SampleItems = ConvertTo-PimGridRow -Row @($script:SampleInventory.Rows)
}

AfterAll {
    Remove-Module AzPimInventory -ErrorAction SilentlyContinue
}

Describe 'ConvertTo-PimGridRow' {
    It 'maps report rows to typed grid rows' {
        $script:SampleItems.Count | Should -Be 10
        $script:SampleItems[0].GetType().Name | Should -Be 'PimGridRow'
    }

    It 'derives privileged flag, sort keys and search text' {
        $permanent = $script:SampleItems | Where-Object { $_.PrincipalName -eq 'Legacy Break Glass Account' }
        $permanent.IsPrivileged | Should -BeTrue
        $permanent.DaysSort | Should -Be ([int]::MaxValue)
        $permanent.ExpirySort | Should -Be '9999'
        $permanent.StatusOrder | Should -Be 5
        $permanent.SearchText | Should -BeLike '*user access administrator*'

        $expired = $script:SampleItems | Where-Object { $_.Status -eq 'Expired' }
        $expired.DaysSort | Should -BeLessThan 0
        $expired.StatusOrder | Should -Be 1
    }

    It 'keeps a reference to the source row for exports' {
        $script:SampleItems[0].Source.PSObject.Properties.Name | Should -Contain 'InstanceId'
    }
}

Describe 'Test-PimGridRowFilter' {
    It 'filters by <Case>' -ForEach @(
        @{ Case = 'nothing'; Params = @{}; Expected = 10 }
        @{ Case = 'status'; Params = @{ Status = 'Expiring' }; Expected = 5 }
        @{ Case = 'kind'; Params = @{ Kind = 'Eligible' }; Expected = 7 }
        @{ Case = 'privileged'; Params = @{ PrivilegedOnly = $true }; Expected = 7 }
        @{ Case = 'standing privileged'; Params = @{ Kind = 'Active'; PrivilegedOnly = $true }; Expected = 3 }
        @{ Case = 'search, case-insensitive'; Params = @{ Term = '  TERRAFORM ' }; Expected = 1 }
        @{ Case = 'search by subscription'; Params = @{ Term = 'pirsa-prod-web' }; Expected = 2 }
        @{ Case = 'no match'; Params = @{ Term = 'zzz' }; Expected = 0 }
    ) {
        $p = $Params
        @($script:SampleItems | Where-Object { Test-PimGridRowFilter -Item $_ @p }).Count | Should -Be $Expected
    }

    It 'rejects null items' {
        Test-PimGridRowFilter -Item $null | Should -BeFalse
    }
}

Describe 'Get-PimSummary' {
    It 'counts the KPI cards' {
        $s = Get-PimSummary -Item $script:SampleItems
        $s.Total | Should -Be 10
        $s.Expired | Should -Be 1
        $s.Expiring | Should -Be 5
        $s.Permanent | Should -Be 3
        $s.Eligible | Should -Be 7
        $s.Active | Should -Be 3
        $s.Standing | Should -Be 3
        $s.Pending | Should -Be 0
    }

    It 'handles an empty inventory' {
        (Get-PimSummary -Item @()).Total | Should -Be 0
    }
}

Describe 'Refresh schedule' {
    BeforeAll {
        $script:From = [datetime]'2026-11-27T13:00:00'
    }

    It 'parses <Text> as <Expected>' -ForEach @(
        @{ Text = '07:00'; Expected = '07:00:00' }
        @{ Text = '7:05'; Expected = '07:05:00' }
        @{ Text = ' 18:30 '; Expected = '18:30:00' }
        @{ Text = '23:59'; Expected = '23:59:00' }
    ) {
        (ConvertTo-TimeOfDay $Text).ToString() | Should -Be $Expected
    }

    It 'rejects invalid time <Text>' -ForEach @(
        @{ Text = '24:00' }, @{ Text = '25:99' }, @{ Text = '7' }, @{ Text = 'noon' }, @{ Text = '' }
    ) {
        ConvertTo-TimeOfDay $Text | Should -BeNullOrEmpty
    }

    It 'returns no next refresh when auto-refresh is off' {
        Get-NextRefreshTime -Schedule '0' -From $script:From | Should -BeNullOrEmpty
    }

    It 'adds the interval for schedule <Schedule>' -ForEach @(
        @{ Schedule = '15'; Minutes = 15 }
        @{ Schedule = '60'; Minutes = 60 }
        @{ Schedule = '720'; Minutes = 720 }
    ) {
        Get-NextRefreshTime -Schedule $Schedule -From $script:From | Should -Be $script:From.AddMinutes($Minutes)
    }

    It 'schedules a daily time later today' {
        Get-NextRefreshTime -Schedule 'daily' -DailyTime '18:30' -From $script:From | Should -Be ([datetime]'2026-11-27T18:30:00')
    }

    It 'rolls a daily time that has passed to tomorrow' {
        Get-NextRefreshTime -Schedule 'daily' -DailyTime '07:00' -From $script:From | Should -Be ([datetime]'2026-11-28T07:00:00')
    }

    It 'rolls a daily time equal to now to tomorrow (no double refresh)' {
        Get-NextRefreshTime -Schedule 'daily' -DailyTime '13:00' -From $script:From | Should -Be ([datetime]'2026-11-28T13:00:00')
    }

    It 'returns nothing for an invalid daily time' {
        Get-NextRefreshTime -Schedule 'daily' -DailyTime '99:99' -From $script:From | Should -BeNullOrEmpty
    }

    It 'describes the schedule for the status bar' {
        Format-ScheduleText -NextRefresh $null -Now $script:From | Should -Be 'Auto-refresh off'
        Format-ScheduleText -NextRefresh $script:From.AddMinutes(30) -Now $script:From | Should -Be 'Next refresh 13:30 (in 30 min)'
        Format-ScheduleText -NextRefresh $script:From.AddMinutes(150) -Now $script:From -LastRefresh $script:From |
            Should -Be 'Last refreshed 13:00   ·   Next refresh 15:30 (in 2h 30m)'
        Format-ScheduleText -NextRefresh ([datetime]'2026-11-28T07:00:00') -Now $script:From | Should -BeLike 'Next refresh * 07:00 (in 18h 00m)'
    }
}

Describe 'App settings' {
    It 'returns defaults when no settings file exists' {
        $s = Read-AppSetting -Path (Join-Path $TestDrive 'missing.json')
        $s.RefreshSchedule | Should -Be '60'
        $s.DailyTime | Should -Be '07:00'
        $s.WarningDays | Should -Be 30
    }

    It 'round-trips saved settings' {
        $path = Join-Path $TestDrive 'roundtrip/settings.json'
        $s = Get-DefaultAppSetting
        $s.Tenant = 'contoso.onmicrosoft.com'; $s.ManagementGroupId = 'mg-root'; $s.RefreshSchedule = 'daily'; $s.DailyTime = '18:30'
        $s.WarningDays = 60; $s.IncludeFuture = $true

        Save-AppSetting -Setting $s -Path $path
        $r = Read-AppSetting -Path $path

        $r.Tenant | Should -Be 'contoso.onmicrosoft.com'
        $r.ManagementGroupId | Should -Be 'mg-root'
        $r.RefreshSchedule | Should -Be 'daily'
        $r.DailyTime | Should -Be '18:30'
        $r.WarningDays | Should -Be 60
        $r.IncludeFuture | Should -BeTrue
    }

    It 'sanitises invalid or tampered values' {
        $path = Join-Path $TestDrive 'bad.json'
        '{ "RefreshSchedule": "5", "DailyTime": "99:00", "WarningDays": 1000 }' | Set-Content -LiteralPath $path
        $r = Read-AppSetting -Path $path
        $r.RefreshSchedule | Should -Be '60'
        $r.DailyTime | Should -Be '07:00'
        $r.WarningDays | Should -Be 30
    }

    It 'survives a corrupt settings file' {
        $path = Join-Path $TestDrive 'corrupt.json'
        'not json {' | Set-Content -LiteralPath $path
        (Read-AppSetting -Path $path).RefreshSchedule | Should -Be '60'
    }
}

Describe 'Get-PortalScopeUrl' {
    It 'links resources to their access control blade' {
        Get-PortalScopeUrl -Scope '/subscriptions/s1/resourceGroups/rg' -TenantId 't1' |
            Should -Be 'https://portal.azure.com/#@t1/resource/subscriptions/s1/resourceGroups/rg/users'
    }

    It 'links management groups to their overview' {
        Get-PortalScopeUrl -Scope '/providers/Microsoft.Management/managementGroups/mg' |
            Should -Be 'https://portal.azure.com/#/resource/providers/Microsoft.Management/managementGroups/mg/overview'
    }
}

Describe 'Main window XAML' {
    It 'defines every control the app code references' {
        $xaml  = Get-Content -LiteralPath $script:XamlFile -Raw
        $code  = Get-Content -LiteralPath $script:AppScript -Raw
        $named = [regex]::Matches($xaml, 'x:Name="(\w+)"') | ForEach-Object { $_.Groups[1].Value }
        $used  = [regex]::Matches($code, '\$(?:script:)?ui\.(\w+)') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique
        $missing = @($used | Where-Object { $_ -notin $named })
        $missing | Should -BeNullOrEmpty
    }

    It 'loads in WPF with all named controls' -Skip:(-not $script:IsWindowsHost) {
        # WPF needs an STA thread; run the load on a dedicated STA runspace.
        $runspace = [runspacefactory]::CreateRunspace()
        $runspace.ApartmentState = 'STA'
        $runspace.ThreadOptions = 'ReuseThread'
        $runspace.Open()
        $ps = [powershell]::Create()
        $ps.Runspace = $runspace
        [void]$ps.AddScript({
                param($AppScript)
                . $AppScript
                $loaded = Import-AppXaml -Path $script:XamlPath
                $loaded.Window.Close()
                [pscustomobject]@{ Count = $loaded.Controls.Count; Names = @($loaded.Controls.Keys) }
            }.ToString()).AddArgument($script:AppScript)
        try {
            $result = $ps.Invoke()
            $ps.Streams.Error | Should -BeNullOrEmpty
            $result[0].Names | Should -Contain 'AssignmentGrid'
            $result[0].Names | Should -Contain 'RefreshScheduleCombo'
            $result[0].Names | Should -Contain 'DailyTimeBox'
            $result[0].Count | Should -BeGreaterThan 50
        }
        finally {
            $ps.Dispose()
            $runspace.Dispose()
        }
    }
}
