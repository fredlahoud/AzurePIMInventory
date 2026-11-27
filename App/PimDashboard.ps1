#Requires -Version 5.1
<#
.SYNOPSIS
    Azure PIM Inventory desktop dashboard.

.DESCRIPTION
    A Windows desktop app (WPF) over the AzPimInventory module:

      * Sign in to any Entra tenant through the browser (MFA / Conditional Access
        supported), or reuse an existing Azure PowerShell sign-in.
      * Pick a management group, see KPI cards and a filterable, sortable grid of
        every eligible and active PIM assignment with expiry status.
      * Auto-refresh on a schedule the user chooses: off, every 15 minutes to every
        12 hours, or daily at a set time. Default: every hour.
      * Export the visible rows to CSV or open them as an HTML report.

    All Azure calls run on background runspaces, so the window stays responsive.
    Settings are kept in %LOCALAPPDATA%\AzPimInventory\settings.json and errors are
    logged to %LOCALAPPDATA%\AzPimInventory\app.log.

    Launch by double-clicking Start-PimDashboard.cmd in the repository root, or run
    Install-DesktopShortcut.ps1 once to add Desktop and Start menu shortcuts.

.PARAMETER Demo
    Start straight into the sample data, without signing in.
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',
    Justification = 'Internal UI helpers that update window state, not cmdlets.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '',
    Justification = 'WPF event and task callbacks have fixed signatures; some parameters are used inside nested scriptblocks.')]
[CmdletBinding()]
param(
    [switch]$Demo
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:AppVersion        = '2.0.0'
$script:AppRoot           = $PSScriptRoot
$script:RepoRoot          = Split-Path -Parent $PSScriptRoot
$script:ModuleManifest    = Join-Path $script:RepoRoot 'src\AzPimInventory\AzPimInventory.psd1'
$script:SampleDataPath    = Join-Path $script:RepoRoot 'sample-data.json'
$script:XamlPath          = Join-Path $script:AppRoot 'MainWindow.xaml'
$script:IconPath          = Join-Path $script:AppRoot 'assets\app.ico'
$script:DemoReferenceDate = [datetime]'2026-09-11T01:36:00'
$script:DataDir           = Join-Path $(if ($env:LOCALAPPDATA) { $env:LOCALAPPDATA } else { [System.IO.Path]::GetTempPath() }) 'AzPimInventory'
$script:SettingsPath      = Join-Path $script:DataDir 'settings.json'
$script:LogPath           = Join-Path $script:DataDir 'app.log'

#region Model -----------------------------------------------------------------

# Strongly typed grid row: WPF sorts and binds typed properties reliably and fast.
class PimGridRow {
    [string]$PrincipalName
    [string]$PrincipalType
    [string]$PrincipalEmail
    [string]$PrincipalId
    [string]$RoleName
    [bool]$IsPrivileged
    [string]$AssignmentKind
    [string]$MemberType
    [string]$SubscriptionName
    [string]$ScopeName
    [string]$ScopeType
    [string]$Scope
    [string]$StartDateUtc
    [string]$ExpiryDateUtc
    [string]$ExpirySort
    [string]$DaysRemaining
    [int]$DaysSort
    [string]$Status
    [int]$StatusOrder
    [string]$InstanceId
    [string]$SearchText
    hidden [object]$Source
}

$script:StatusOrder = @{ 'Unknown' = 0; 'Expired' = 1; 'Expiring' = 2; 'Pending' = 3; 'Active' = 4; 'Permanent' = 5 }

function ConvertTo-PimGridRow {
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Row)

    $list = [System.Collections.Generic.List[object]]::new()
    foreach ($r in $Row) {
        $days  = [string]$r.DaysRemaining
        $item  = [PimGridRow]::new()
        $item.PrincipalName    = [string]$r.PrincipalName
        $item.PrincipalType    = [string]$r.PrincipalType
        $item.PrincipalEmail   = [string]$r.PrincipalEmail
        $item.PrincipalId      = [string]$r.PrincipalId
        $item.RoleName         = [string]$r.RoleName
        $item.IsPrivileged     = ([string]$r.PrivilegedRole -eq 'Yes')
        $item.AssignmentKind   = [string]$r.AssignmentKind
        $item.MemberType       = [string]$r.MemberType
        $item.SubscriptionName = [string]$r.SubscriptionName
        $item.ScopeName        = [string]$r.ScopeName
        $item.ScopeType        = [string]$r.ScopeType
        $item.Scope            = [string]$r.Scope
        $item.StartDateUtc     = [string]$r.StartDateUtc
        $item.ExpiryDateUtc    = [string]$r.ExpiryDateUtc
        $item.ExpirySort       = if ($item.ExpiryDateUtc) { $item.ExpiryDateUtc } else { '9999' }
        $item.DaysRemaining    = $days
        $item.DaysSort         = if ([string]::IsNullOrEmpty($days)) { [int]::MaxValue } else { [int]$days }
        $item.Status           = [string]$r.Status
        $item.StatusOrder      = if ($script:StatusOrder.ContainsKey($item.Status)) { $script:StatusOrder[$item.Status] } else { 9 }
        $item.InstanceId       = [string]$r.InstanceId
        $item.SearchText       = (@($item.PrincipalName, $item.PrincipalEmail, $item.PrincipalId, $item.RoleName,
                $item.SubscriptionName, $item.ScopeName, $item.Scope, $item.AssignmentKind, $item.Status) -join ' ').ToLowerInvariant()
        $item.Source           = $r
        $list.Add($item)
    }
    return , $list
}

function Test-PimGridRowFilter {
    param(
        $Item,
        [string]$Term = '',
        [string]$Status = '',
        [string]$Kind = '',
        [bool]$PrivilegedOnly = $false
    )

    if ($null -eq $Item) { return $false }
    if ($Status -and $Item.Status -ne $Status) { return $false }
    if ($Kind -and $Item.AssignmentKind -ne $Kind) { return $false }
    if ($PrivilegedOnly -and -not $Item.IsPrivileged) { return $false }
    if ($Term -and $Item.SearchText.IndexOf($Term.Trim().ToLowerInvariant(), [System.StringComparison]::Ordinal) -lt 0) { return $false }
    return $true
}

function Get-PimSummary {
    param([AllowEmptyCollection()][object[]]$Item = @())

    $s = [ordered]@{ Total = 0; Expired = 0; Expiring = 0; Permanent = 0; Pending = 0; Unknown = 0; Eligible = 0; Active = 0; Standing = 0 }
    foreach ($i in $Item) {
        $s.Total++
        switch ($i.Status) {
            'Expired'   { $s.Expired++ }
            'Expiring'  { $s.Expiring++ }
            'Permanent' { $s.Permanent++ }
            'Pending'   { $s.Pending++ }
            'Unknown'   { $s.Unknown++ }
        }
        if ($i.AssignmentKind -eq 'Eligible') { $s.Eligible++ }
        if ($i.AssignmentKind -eq 'Active') {
            $s.Active++
            if ($i.IsPrivileged) { $s.Standing++ }
        }
    }
    return $s
}

#endregion

#region Schedule ---------------------------------------------------------------

<#
    Parses a 24-hour time of day such as '7:00', '07:00' or '18:30'.
    Returns a TimeSpan, or $null when the text is not a valid time.
#>
function ConvertTo-TimeOfDay {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $parsed = [datetime]::MinValue
    $formats = [string[]]@('H:mm', 'HH:mm')
    if ([datetime]::TryParseExact($Text.Trim(), $formats, [cultureinfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
        return $parsed.TimeOfDay
    }
    return $null
}

<#
    Works out when the next automatic refresh is due.
      Schedule '0'      -> $null (auto-refresh off)
      Schedule '<n>'    -> From + n minutes
      Schedule 'daily'  -> the next occurrence of DailyTime (local time) after From
#>
function Get-NextRefreshTime {
    param(
        [string]$Schedule,
        [string]$DailyTime = '07:00',
        [datetime]$From = (Get-Date)
    )

    if ($Schedule -eq 'daily') {
        $time = ConvertTo-TimeOfDay $DailyTime
        if ($null -eq $time) { return $null }
        $candidate = $From.Date.Add($time)
        if ($candidate -le $From) { $candidate = $candidate.AddDays(1) }
        return $candidate
    }

    $minutes = 0
    if ([int]::TryParse($Schedule, [ref]$minutes) -and $minutes -gt 0) {
        return $From.AddMinutes($minutes)
    }
    return $null
}

function Format-ScheduleText {
    param($NextRefresh, [datetime]$Now = (Get-Date), $LastRefresh)

    $parts = @()
    if ($LastRefresh) { $parts += 'Last refreshed ' + ([datetime]$LastRefresh).ToString('HH:mm') }
    if ($null -eq $NextRefresh) {
        $parts += 'Auto-refresh off'
    }
    else {
        $next = [datetime]$NextRefresh
        $span = $next - $Now
        $when = if ($next.Date -eq $Now.Date) { $next.ToString('HH:mm') } else { $next.ToString('ddd HH:mm') }
        $in   = if ($span.TotalMinutes -lt 1) { 'due now' }
                elseif ($span.TotalMinutes -lt 60) { 'in {0} min' -f [math]::Ceiling($span.TotalMinutes) }
                else { 'in {0}h {1:00}m' -f [math]::Floor($span.TotalHours), $span.Minutes }
        $parts += "Next refresh $when ($in)"
    }
    return ($parts -join '   ·   ')
}

#endregion

#region Settings and logging ----------------------------------------------------

function Get-DefaultAppSetting {
    return [ordered]@{
        Tenant            = ''
        ManagementGroupId = ''
        WarningDays       = 30
        IncludeFuture     = $false
        IncludeJit        = $false
        RefreshSchedule   = '60'
        DailyTime         = '07:00'
    }
}

function Read-AppSetting {
    param([string]$Path = $script:SettingsPath)

    $settings = Get-DefaultAppSetting
    if (Test-Path -LiteralPath $Path) {
        try {
            $json = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
            foreach ($key in @($settings.Keys)) {
                if ($json.PSObject.Properties.Name -contains $key -and $null -ne $json.$key) {
                    $settings[$key] = $json.$key
                }
            }
        }
        catch {
            Write-AppLog "Settings file unreadable, defaults used: $($_.Exception.Message)"
        }
    }

    # Validate everything that came from disk.
    $settings.RefreshSchedule = [string]$settings.RefreshSchedule
    if ($settings.RefreshSchedule -notin @('0', '15', '30', '60', '120', '240', '480', '720', 'daily')) { $settings.RefreshSchedule = '60' }
    if ($null -eq (ConvertTo-TimeOfDay ([string]$settings.DailyTime))) { $settings.DailyTime = '07:00' }
    if ([int]$settings.WarningDays -notin @(7, 14, 30, 60, 90)) { $settings.WarningDays = 30 }
    $settings.IncludeFuture = [bool]$settings.IncludeFuture
    $settings.IncludeJit    = [bool]$settings.IncludeJit
    $settings.Tenant            = [string]$settings.Tenant
    $settings.ManagementGroupId = [string]$settings.ManagementGroupId
    return $settings
}

function Save-AppSetting {
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$Setting,
        [string]$Path = $script:SettingsPath
    )
    try {
        $dir = Split-Path -Parent $Path
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        [pscustomobject]$Setting | ConvertTo-Json | Set-Content -LiteralPath $Path -Encoding UTF8
    }
    catch {
        Write-AppLog "Could not save settings: $($_.Exception.Message)"
    }
}

function Write-AppLog {
    param([string]$Message)
    try {
        if (-not (Test-Path -LiteralPath $script:DataDir)) { New-Item -ItemType Directory -Path $script:DataDir -Force | Out-Null }
        $line = '{0:yyyy-MM-dd HH:mm:ss}  {1}' -f (Get-Date), $Message
        Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8
    }
    catch { Write-Verbose "Log write failed: $($_.Exception.Message)" }
}

function Get-PortalScopeUrl {
    param([string]$Scope, [string]$TenantId = '')

    if ([string]::IsNullOrWhiteSpace($Scope)) { return $null }
    $tenantPart = if ($TenantId) { "@$TenantId" } else { '' }
    if ($Scope -match '^/providers/Microsoft\.Management/managementGroups/([^/]+)$') {
        return "https://portal.azure.com/#$tenantPart/resource$Scope/overview"
    }
    # The 'users' blade is Access control (IAM) for subscriptions, resource groups and resources.
    return "https://portal.azure.com/#$tenantPart/resource$Scope/users"
}

#endregion

#region Background work ------------------------------------------------------

$script:Tasks = [System.Collections.Generic.List[object]]::new()

<#
    Runs a script on a fresh background runspace. OnComplete runs on the UI thread
    (from the poll timer) with the output and any terminating error.
#>
function Start-AppTask {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][scriptblock]$Script,
        [hashtable]$Parameters = @{},
        [Parameter(Mandatory)][scriptblock]$OnComplete
    )

    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.Open()
    $ps = [powershell]::Create()
    $ps.Runspace = $runspace
    # Pass the script as text: a scriptblock object stays bound to the UI runspace.
    [void]$ps.AddScript($Script.ToString())
    if ($Parameters.Count -gt 0) { [void]$ps.AddParameters($Parameters) }

    $task = [pscustomobject]@{
        Name       = $Name
        PowerShell = $ps
        Runspace   = $runspace
        Handle     = $ps.BeginInvoke()
        OnComplete = $OnComplete
        Started    = Get-Date
        Cancelled  = $false
    }
    $script:Tasks.Add($task)
    return $task
}

function Stop-AppTask {
    param([Parameter(Mandatory)]$Task)
    $Task.Cancelled = $true
    try { [void]$Task.PowerShell.BeginStop($null, $null) } catch { Write-AppLog "Stop failed: $($_.Exception.Message)" }
}

function Receive-AppTask {
    foreach ($task in @($script:Tasks)) {
        if (-not $task.Handle.IsCompleted) { continue }
        [void]$script:Tasks.Remove($task)

        $output = $null
        $failure = $null
        try {
            $output = @($task.PowerShell.EndInvoke($task.Handle))
        }
        catch {
            $failure = if ($_.Exception.InnerException) { $_.Exception.InnerException } else { $_.Exception }
        }
        finally {
            $task.PowerShell.Dispose()
            $task.Runspace.Dispose()
        }

        if ($task.Cancelled) { $failure = [System.OperationCanceledException]::new('Cancelled.') }
        if ($failure) { Write-AppLog "$($task.Name) failed: $($failure.Message)" }

        try { & $task.OnComplete $output $failure }
        catch { Show-AppError -Message "$($task.Name): $($_.Exception.Message)" }
    }
}

# Script bodies for background runspaces. They must be self-contained.

$script:CheckSessionScript = {
    $ErrorActionPreference = 'Stop'
    if (-not (Get-Module -ListAvailable -Name Az.Accounts)) {
        return [pscustomobject]@{ AzAvailable = $false; Account = $null; TenantId = $null; Context = $null }
    }
    Import-Module Az.Accounts -WarningAction SilentlyContinue
    $ctx = Get-AzContext -ErrorAction SilentlyContinue
    [pscustomobject]@{
        AzAvailable = $true
        Account     = if ($ctx) { $ctx.Account.Id } else { $null }
        TenantId    = if ($ctx) { $ctx.Tenant.Id } else { $null }
        Context     = $ctx
    }
}

$script:SignInScript = {
    param([string]$Tenant)
    $ErrorActionPreference = 'Stop'
    Import-Module Az.Accounts -WarningAction SilentlyContinue
    # Process-scoped only: never changes the user's saved Az configuration.
    # Browser sign-in works from a windowless background thread; the newer
    # interactive subscription picker would wait on a console nobody can see.
    try { Update-AzConfig -EnableLoginByWam $false -Scope Process -WarningAction SilentlyContinue | Out-Null } catch { $null = $_ }
    try { Update-AzConfig -LoginExperienceV2 Off -Scope Process -WarningAction SilentlyContinue | Out-Null } catch { $null = $_ }
    $params = @{ WarningAction = 'SilentlyContinue'; ErrorAction = 'Stop' }
    if (-not [string]::IsNullOrWhiteSpace($Tenant)) { $params.Tenant = $Tenant.Trim() }
    $ctx = (Connect-AzAccount @params).Context
    [pscustomobject]@{ Account = $ctx.Account.Id; TenantId = $ctx.Tenant.Id; Context = $ctx }
}

$script:SignOutScript = {
    $ErrorActionPreference = 'Stop'
    Import-Module Az.Accounts -WarningAction SilentlyContinue
    Disconnect-AzAccount -WarningAction SilentlyContinue | Out-Null
}

$script:ListManagementGroupsScript = {
    param([string]$ModulePath, $Context)
    $ErrorActionPreference = 'Stop'
    Import-Module Az.Accounts -WarningAction SilentlyContinue
    if (-not (Get-AzContext -ErrorAction SilentlyContinue) -and $null -ne $Context) { Set-AzContext -Context $Context | Out-Null }
    Import-Module $ModulePath -Force
    Get-PimManagementGroup
}

$script:CollectScript = {
    param([string]$ModulePath, $Context, [hashtable]$Options, [hashtable]$Progress)
    $ErrorActionPreference = 'Stop'
    if ($Options.ContainsKey('RootManagementGroupId')) {
        Import-Module Az.Accounts -WarningAction SilentlyContinue
        if (-not (Get-AzContext -ErrorAction SilentlyContinue) -and $null -ne $Context) { Set-AzContext -Context $Context | Out-Null }
    }
    Import-Module $ModulePath -Force
    Get-PimInventory @Options -ProgressState $Progress -WarningAction SilentlyContinue -InformationAction SilentlyContinue
}

$script:InstallAzScript = {
    $ErrorActionPreference = 'Stop'
    Install-Module Az.Accounts -Scope CurrentUser -Repository PSGallery -Force -AllowClobber
}

#endregion

#region UI --------------------------------------------------------------------

function Import-AppXaml {
    param([Parameter(Mandatory)][string]$Path)

    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
    # Explicit UTF-8: Windows PowerShell 5.1 would otherwise read it as ANSI.
    $text   = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    [xml]$xml = $text
    $reader = [System.Xml.XmlNodeReader]::new($xml)
    $window = [System.Windows.Markup.XamlReader]::Load($reader)

    $controls = @{}
    foreach ($m in [regex]::Matches($text, 'x:Name="([A-Za-z0-9_]+)"')) {
        $name = $m.Groups[1].Value
        $element = $window.FindName($name)
        if ($null -ne $element) { $controls[$name] = $element }
    }
    return [pscustomobject]@{ Window = $window; Controls = $controls }
}

function Get-IconImage {
    param([string]$Path, [int]$PreferredSize = 256)
    try {
        $decoder = [System.Windows.Media.Imaging.IconBitmapDecoder]::new([uri]$Path,
            [System.Windows.Media.Imaging.BitmapCreateOptions]::None,
            [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad)
        $frames = @($decoder.Frames | Sort-Object { [math]::Abs($_.PixelWidth - $PreferredSize) })
        return $frames[0]
    }
    catch {
        Write-AppLog "Icon load failed: $($_.Exception.Message)"
        return $null
    }
}

function Show-AppError {
    param([string]$Message)
    Write-AppLog "ERROR: $Message"
    if ($script:ui) {
        Set-Banner -Text $Message -Kind Error
        Set-Status $Message
    }
}

function Invoke-UiAction {
    param([Parameter(Mandatory)][scriptblock]$Action)
    try { & $Action }
    catch { Show-AppError -Message $_.Exception.Message }
}

function Set-Status {
    param([string]$Text)
    $script:ui.StatusText.Text = $Text
}

function Set-Banner {
    param(
        [string]$Text,
        [ValidateSet('Warning', 'Error', 'Info')][string]$Kind = 'Warning',
        [string]$ActionText,
        [scriptblock]$Action
    )

    $b = $script:ui.Banner
    if ([string]::IsNullOrEmpty($Text)) {
        $b.Visibility = 'Collapsed'
        return
    }
    $colors = switch ($Kind) {
        'Error' { @('#FDE7E9', '#F1BBBC', '#5C0F12') }
        'Info'  { @('#DEECF9', '#A9CDEF', '#0B3A63') }
        default { @('#FFF4CE', '#E9C96A', '#3B2E00') }
    }
    $converter = [System.Windows.Media.BrushConverter]::new()
    $b.Background  = $converter.ConvertFromString($colors[0])
    $b.BorderBrush = $converter.ConvertFromString($colors[1])
    $script:ui.BannerText.Foreground = $converter.ConvertFromString($colors[2])
    $script:ui.BannerText.Text = $Text
    $script:BannerAction = $Action
    if ($ActionText -and $Action) {
        $script:ui.BannerActionButton.Content = $ActionText
        $script:ui.BannerActionButton.Visibility = 'Visible'
    }
    else {
        $script:ui.BannerActionButton.Visibility = 'Collapsed'
    }
    $b.Visibility = 'Visible'
}

function Get-ComboTag {
    param($Combo)
    $item = $Combo.SelectedItem
    if ($null -eq $item) { return '' }
    return [string]$item.Tag
}

function Select-ComboTag {
    param($Combo, [string]$Tag)
    foreach ($item in $Combo.Items) {
        if ([string]$item.Tag -eq $Tag) { $Combo.SelectedItem = $item; return }
    }
}

function Get-SelectedManagementGroupId {
    $combo = $script:ui.ManagementGroupCombo
    $selected = $combo.SelectedItem
    if ($null -ne $selected -and $combo.Text -eq $selected.Label) { return [string]$selected.Id }
    $text = [string]$combo.Text
    # Accept "Display name (id)" pasted back in, as well as a bare id.
    if ($text -match '\(([^()]+)\)\s*$') { return $Matches[1].Trim() }
    return $text.Trim()
}

function Set-Busy {
    param([bool]$Busy, [string]$Message = '')

    $script:State.Busy = $Busy
    $ui = $script:ui
    $ui.BusyBar.Visibility = if ($Busy) { 'Visible' } else { 'Collapsed' }
    $ui.BusyBar.IsIndeterminate = $true
    $ui.RefreshButton.Content = if ($Busy) { 'Refreshing…' } else { 'Refresh now' }
    Update-CommandState
    if ($Message) { Set-Status $Message }
}

function Update-CommandState {
    $ui = $script:ui
    $mode = $script:State.Mode
    $hasData = $null -ne $script:State.Inventory
    $ui.RefreshButton.IsEnabled    = ($mode -ne 'SignedOut') -and -not $script:State.Busy
    $ui.ExportCsvButton.IsEnabled  = $hasData
    $ui.HtmlReportButton.IsEnabled = $hasData
    $ui.ManagementGroupCombo.IsEnabled = ($mode -eq 'SignedIn')
}

function Update-HeaderState {
    $ui = $script:ui
    switch ($script:State.Mode) {
        'SignedIn' {
            $ui.AccountText.Text = $script:State.Account
            $ui.TenantText.Text  = "Tenant $($script:State.TenantId)"
            $ui.AccountPanel.Visibility     = 'Visible'
            $ui.TenantInputPanel.Visibility = 'Collapsed'
            $ui.SignInButton.Visibility     = 'Collapsed'
            $ui.SignOutButton.Visibility    = 'Visible'
            $ui.SignOutButton.Content       = 'Sign out'
            $ui.SignOutButton.ToolTip       = 'Signs this account out of Azure PowerShell on this computer.'
        }
        'Demo' {
            $ui.AccountText.Text = 'Sample data'
            $ui.TenantText.Text  = 'Demo mode - not connected to Azure'
            $ui.AccountPanel.Visibility     = 'Visible'
            $ui.TenantInputPanel.Visibility = 'Collapsed'
            $ui.SignInButton.Visibility     = 'Collapsed'
            $ui.SignOutButton.Visibility    = 'Visible'
            $ui.SignOutButton.Content       = 'Exit demo'
            $ui.SignOutButton.ToolTip       = $null
        }
        default {
            $ui.AccountPanel.Visibility     = 'Collapsed'
            $ui.TenantInputPanel.Visibility = 'Visible'
            $ui.SignInButton.Visibility     = 'Visible'
            $ui.SignOutButton.Visibility    = 'Collapsed'
        }
    }
    Update-CommandState
}

function Show-WelcomePanel {
    param([string]$Title, [string]$Body, [bool]$ShowActions = $true)
    $ui = $script:ui
    if ($Title) { $ui.WelcomeTitle.Text = $Title }
    if ($Body) { $ui.WelcomeBody.Text = $Body }
    $visibility = if ($ShowActions) { 'Visible' } else { 'Collapsed' }
    $ui.WelcomeSignInButton.Visibility = $visibility
    $ui.DemoButton.Visibility = $visibility
    $ui.WelcomePanel.Visibility = 'Visible'
}

function Update-Summary {
    $ui = $script:ui
    $all = if ($script:State.GridItems) { @($script:State.GridItems) } else { @() }
    $s = Get-PimSummary -Item $all
    $ui.TotalCount.Text     = $s.Total
    $ui.ExpiredCount.Text   = $s.Expired
    $ui.ExpiringCount.Text  = $s.Expiring
    $ui.PermanentCount.Text = $s.Permanent
    $ui.EligibleCount.Text  = $s.Eligible
    $ui.ActiveCount.Text    = $s.Active
    $ui.StandingCount.Text  = $s.Standing
    $ui.PendingCount.Text   = $s.Pending
    $ui.CardPending.Visibility = if ($s.Pending -gt 0) { 'Visible' } else { 'Collapsed' }
    $days = if ($script:State.Inventory) { $script:State.Inventory.ExpiryWarningDays } else { Get-ComboTag $ui.WarningDaysCombo }
    $ui.ExpiringLabel.Text = "Expiring ≤ $days days"
}

function Update-FilterState {
    $ui = $script:ui
    $script:FilterState = @{
        Term           = [string]$ui.SearchBox.Text
        Status         = Get-ComboTag $ui.StatusFilterCombo
        Kind           = Get-ComboTag $ui.KindFilterCombo
        PrivilegedOnly = [bool]$ui.PrivilegedOnlyCheck.IsChecked
    }
}

function Update-GridFilter {
    Update-FilterState
    $view = $script:State.View
    if ($null -eq $view) { $script:ui.CountText.Text = ''; return }
    $view.Refresh()
    $script:ui.CountText.Text = "Showing $($view.Count) of $(@($script:State.GridItems).Count)"
}

function Set-DashboardInventory {
    param([Parameter(Mandatory)]$Inventory)

    $ui = $script:ui
    $items = ConvertTo-PimGridRow -Row @($Inventory.Rows)

    # Keep the user's sort order across refreshes.
    $previousSort = @()
    if ($script:State.View) { $previousSort = @($script:State.View.SortDescriptions) }

    $view = [System.Windows.Data.ListCollectionView]::new($items)
    $view.Filter = [Predicate[object]] {
        param($row)
        $f = $script:FilterState
        Test-PimGridRowFilter -Item $row -Term $f.Term -Status $f.Status -Kind $f.Kind -PrivilegedOnly $f.PrivilegedOnly
    }
    if ($previousSort.Count -eq 0) {
        $previousSort = @(
            [System.ComponentModel.SortDescription]::new('StatusOrder', 'Ascending')
            [System.ComponentModel.SortDescription]::new('DaysSort', 'Ascending')
        )
    }

    $script:State.Inventory   = $Inventory
    $script:State.GridItems   = $items
    $script:State.View        = $view
    $script:State.LastRefresh = Get-Date
    $ui.AssignmentGrid.ItemsSource = $view

    # DataGrid clears SortDescriptions when ItemsSource changes, so apply them after
    # assigning, and keep the column header arrows in step.
    $view.SortDescriptions.Clear()
    foreach ($sd in $previousSort) { $view.SortDescriptions.Add($sd) }
    foreach ($column in $ui.AssignmentGrid.Columns) {
        $match = $previousSort | Where-Object { $_.PropertyName -eq $column.SortMemberPath } | Select-Object -First 1
        $column.SortDirection = if ($match) { $match.Direction } else { $null }
    }
    $ui.WelcomePanel.Visibility = 'Collapsed'

    Update-Summary
    Update-GridFilter
    Update-CommandState

    $source = if ($Inventory.Source -eq 'Offline') { "sample data, evaluated as of $($Inventory.GeneratedUtc.ToString('dd MMM yyyy')) UTC" }
              else { "$($Inventory.RootManagementGroupId) - $($Inventory.ManagementGroupCount) management group(s), $($Inventory.SubscriptionCount) subscription(s)" }
    $ui.SubtitleText.Text = "Showing $source"

    $errors = @($Inventory.ScopeErrors)
    if ($errors.Count -gt 0) {
        $detail = ($errors | Select-Object -First 3 | ForEach-Object { "$($_.Scope) ($($_.Kind))" }) -join '; '
        $more = if ($errors.Count -gt 3) { " and $($errors.Count - 3) more" } else { '' }
        Set-Banner -Kind Warning -Text "$($errors.Count) scope(s) could not be read, so this inventory is incomplete: $detail$more. Grant Reader (or User Access Administrator) at the management group and refresh."
    }
    else {
        Set-Banner -Text ''
    }
}

function Clear-Dashboard {
    $ui = $script:ui
    $script:State.Inventory = $null
    $script:State.GridItems = $null
    $script:State.View      = $null
    $ui.AssignmentGrid.ItemsSource = $null
    foreach ($n in 'TotalCount', 'ExpiredCount', 'ExpiringCount', 'PermanentCount', 'EligibleCount', 'ActiveCount', 'StandingCount', 'PendingCount') {
        $ui[$n].Text = '–'
    }
    $ui.CountText.Text = ''
    $ui.SubtitleText.Text = 'Eligible and active role assignments across your management group hierarchy'
    Update-CommandState
}

function Update-Schedule {
    $ui = $script:ui
    $schedule = Get-ComboTag $ui.RefreshScheduleCombo
    $ui.DailyTimePanel.Visibility = if ($schedule -eq 'daily') { 'Visible' } else { 'Collapsed' }

    if ($schedule -eq 'daily' -and $null -eq (ConvertTo-TimeOfDay $ui.DailyTimeBox.Text)) {
        $ui.DailyTimeBox.BorderBrush = [System.Windows.Media.Brushes]::IndianRed
        $ui.DailyTimeBox.ToolTip = 'Enter a 24-hour time such as 07:00 or 18:30.'
        $script:State.NextRefresh = $null
    }
    else {
        $ui.DailyTimeBox.ClearValue([System.Windows.Controls.Control]::BorderBrushProperty)
        $ui.DailyTimeBox.ToolTip = 'Local time of day to refresh, as HH:mm (for example 07:00 or 18:30).'
        $script:State.NextRefresh = if ($script:State.Mode -eq 'SignedOut') { $null }
                                    else { Get-NextRefreshTime -Schedule $schedule -DailyTime $ui.DailyTimeBox.Text }
    }
    Update-ScheduleText
}

function Update-ScheduleText {
    $text = if ($script:State.Mode -eq 'SignedOut') { '' }
            else { Format-ScheduleText -NextRefresh $script:State.NextRefresh -LastRefresh $script:State.LastRefresh }
    $script:ui.ScheduleText.Text = $text
}

function Get-CurrentSetting {
    $ui = $script:ui
    $mgId = if ($script:State.Mode -eq 'SignedIn') { Get-SelectedManagementGroupId } else { $script:Settings.ManagementGroupId }
    return [ordered]@{
        Tenant            = if ($script:State.Mode -eq 'SignedIn') { [string]$script:State.TenantId } else { [string]$ui.TenantBox.Text }
        ManagementGroupId = [string]$mgId
        WarningDays       = [int](Get-ComboTag $ui.WarningDaysCombo)
        IncludeFuture     = [bool]$ui.IncludeFutureCheck.IsChecked
        IncludeJit        = [bool]$ui.IncludeJitCheck.IsChecked
        RefreshSchedule   = Get-ComboTag $ui.RefreshScheduleCombo
        DailyTime         = [string]$ui.DailyTimeBox.Text
    }
}

function Save-CurrentSetting {
    $script:Settings = Get-CurrentSetting
    Save-AppSetting -Setting $script:Settings
}

#endregion

#region Workflows -------------------------------------------------------------

function Start-InventoryRefresh {
    param([string]$Reason = 'Manual')

    if ($script:State.Mode -eq 'SignedOut') { return }
    if ($script:State.Busy) { $script:State.RefreshQueued = $true; return }

    $ui = $script:ui
    $options = @{
        ExpiryWarningDays        = [int](Get-ComboTag $ui.WarningDaysCombo)
        IncludeFutureAssignments = [bool]$ui.IncludeFutureCheck.IsChecked
        IncludeActivations       = [bool]$ui.IncludeJitCheck.IsChecked
    }

    if ($script:State.Mode -eq 'Demo') {
        $options.OfflineSampleData = $script:SampleDataPath
        $options.ReferenceDateUtc  = $script:DemoReferenceDate
    }
    else {
        $mgId = Get-SelectedManagementGroupId
        if ([string]::IsNullOrWhiteSpace($mgId)) {
            Set-Banner -Kind Info -Text 'Choose a management group (or type its ID) and click Refresh now.'
            return
        }
        $options.RootManagementGroupId = $mgId
    }

    $script:State.RefreshQueued = $false
    $script:Progress = [hashtable]::Synchronized(@{ Message = 'Starting'; Percent = -1 })
    Set-Busy -Busy $true -Message $(if ($Reason -eq 'Scheduled') { 'Scheduled refresh started…' } else { 'Collecting PIM assignments…' })
    Write-AppLog "Refresh ($Reason) started: $(($options.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ', ')"

    $null = Start-AppTask -Name 'Refresh' -Script $script:CollectScript -Parameters @{
        ModulePath = $script:ModuleManifest
        Context    = $script:State.Context
        Options    = $options
        Progress   = $script:Progress
    } -OnComplete {
        param($output, $failure)
        $script:Progress = $null
        Set-Busy -Busy $false
        if ($failure) {
            Show-AppError -Message "Refresh failed: $($failure.Message)"
        }
        else {
            $inventory = @($output | Where-Object { $_.PSObject.TypeNames -contains 'AzPimInventory.Result' }) | Select-Object -First 1
            if ($null -eq $inventory) { Show-AppError -Message 'Refresh returned no data.'; return }
            Set-DashboardInventory -Inventory $inventory
            Set-Status ("Loaded {0} assignment(s) at {1:HH:mm}." -f @($inventory.Rows).Count, (Get-Date))
            Write-AppLog "Refresh completed: $(@($inventory.Rows).Count) row(s), $(@($inventory.ScopeErrors).Count) scope error(s)."
        }
        Update-Schedule
        if ($script:State.RefreshQueued) { Start-InventoryRefresh -Reason 'Queued' }
    }
}

function Enter-SignedInState {
    param([string]$Account, [string]$TenantId, $Context)

    $script:State.Mode     = 'SignedIn'
    $script:State.Account  = $Account
    $script:State.TenantId = $TenantId
    $script:State.Context  = $Context
    Update-HeaderState
    Clear-Dashboard
    Set-Banner -Text ''
    Show-WelcomePanel -Title 'Loading management groups…' -Body "Signed in as $Account." -ShowActions $false
    Set-Busy -Busy $true -Message 'Listing management groups…'
    Write-AppLog "Signed in as $Account (tenant $TenantId)."

    $null = Start-AppTask -Name 'ListManagementGroups' -Script $script:ListManagementGroupsScript -Parameters @{
        ModulePath = $script:ModuleManifest
        Context    = $Context
    } -OnComplete {
        param($output, $failure)
        Set-Busy -Busy $false
        $ui = $script:ui
        $groups = @($output | Where-Object { $null -ne $_ -and $_.PSObject.Properties.Name -contains 'Id' })
        $ui.ManagementGroupCombo.ItemsSource = $groups

        if ($failure -or $groups.Count -eq 0) {
            $why = if ($failure) { $failure.Message } else { 'no management groups are visible to this account' }
            Show-WelcomePanel -Title 'Choose a management group' -Body 'Type the management group ID (the name, not the display name) in the box above, then click Refresh now.' -ShowActions $false
            Set-Banner -Kind Warning -Text "Could not list management groups ($why). You can still type a management group ID above."
            if ($script:Settings.ManagementGroupId) { $ui.ManagementGroupCombo.Text = $script:Settings.ManagementGroupId }
            Update-Schedule
            return
        }

        $preferred = $groups | Where-Object { $_.Id -eq $script:Settings.ManagementGroupId } | Select-Object -First 1
        if (-not $preferred) { $preferred = $groups | Where-Object { $_.IsTenantRoot } | Select-Object -First 1 }
        if (-not $preferred) { $preferred = $groups[0] }

        $script:SuppressAutoRefresh = $true
        $ui.ManagementGroupCombo.SelectedItem = $preferred
        $script:SuppressAutoRefresh = $false
        Set-Status "$($groups.Count) management group(s) available."
        Start-InventoryRefresh -Reason 'Sign-in'
    }
}

function Start-SignIn {
    if ($script:State.SignInTask) {
        Stop-AppTask -Task $script:State.SignInTask
        return
    }
    if (-not $script:State.AzAvailable) {
        Show-AzMissingBanner
        return
    }

    $tenant = [string]$script:ui.TenantBox.Text
    $script:ui.SignInButton.Content = 'Cancel sign-in'
    $script:ui.WelcomeSignInButton.Content = 'Cancel sign-in'
    Set-Busy -Busy $true -Message 'Complete the sign-in in your browser…'
    Set-Banner -Kind Info -Text 'A browser window has opened for Microsoft sign-in. Complete it there, then return here.'

    $script:State.SignInTask = Start-AppTask -Name 'SignIn' -Script $script:SignInScript -Parameters @{ Tenant = $tenant } -OnComplete {
        param($output, $failure)
        $script:State.SignInTask = $null
        $script:ui.SignInButton.Content = 'Sign in'
        $script:ui.WelcomeSignInButton.Content = 'Sign in'
        Set-Busy -Busy $false
        if ($failure -is [System.OperationCanceledException] -or ($failure -and $failure.GetType().Name -eq 'PipelineStoppedException')) {
            Set-Banner -Text ''
            Set-Status 'Sign-in cancelled.'
            return
        }
        if ($failure) {
            Show-AppError -Message "Sign-in failed: $($failure.Message)"
            return
        }
        $result = @($output) | Select-Object -Last 1
        Enter-SignedInState -Account $result.Account -TenantId $result.TenantId -Context $result.Context
    }
}

function Start-SignOut {
    if ($script:State.Mode -eq 'Demo') {
        Exit-ToSignedOut
        return
    }
    Set-Busy -Busy $true -Message 'Signing out…'
    $null = Start-AppTask -Name 'SignOut' -Script $script:SignOutScript -OnComplete {
        param($output, $failure)
        Set-Busy -Busy $false
        if ($failure) { Write-AppLog "Sign-out warning: $($failure.Message)" }
        Exit-ToSignedOut
        Set-Status 'Signed out.'
    }
}

function Exit-ToSignedOut {
    $script:State.Mode = 'SignedOut'
    $script:State.Account = $null
    $script:State.TenantId = $null
    $script:State.Context = $null
    $script:State.LastRefresh = $null
    $script:ui.ManagementGroupCombo.ItemsSource = $null
    Clear-Dashboard
    Set-Banner -Text ''
    Update-HeaderState
    Show-WelcomePanel -Title 'Sign in to your Azure tenant' -Body $script:WelcomeBodyText -ShowActions $true
    Update-Schedule
}

function Start-Demo {
    $script:State.Mode = 'Demo'
    $script:State.Context = $null
    Update-HeaderState
    Set-Banner -Text ''
    Update-Schedule
    Start-InventoryRefresh -Reason 'Demo'
}

function Show-AzMissingBanner {
    Set-Banner -Kind Warning -Text 'The Az.Accounts PowerShell module is required to sign in. Install it for your user account (no admin rights needed), or explore the sample data.' `
        -ActionText 'Install Az.Accounts' -Action {
        Set-Busy -Busy $true -Message 'Installing Az.Accounts from the PowerShell Gallery…'
        $null = Start-AppTask -Name 'InstallAz' -Script $script:InstallAzScript -OnComplete {
            param($output, $failure)
            Set-Busy -Busy $false
            if ($failure) { Show-AppError -Message "Installing Az.Accounts failed: $($failure.Message)"; return }
            $script:State.AzAvailable = $true
            Set-Banner -Kind Info -Text 'Az.Accounts installed. You can sign in now.'
            Set-Status 'Az.Accounts installed.'
        }
    }
}

function Start-SessionCheck {
    Set-Busy -Busy $true -Message 'Checking for an existing Azure sign-in…'
    $null = Start-AppTask -Name 'CheckSession' -Script $script:CheckSessionScript -OnComplete {
        param($output, $failure)
        Set-Busy -Busy $false
        $result = @($output) | Select-Object -Last 1
        if ($failure -or $null -eq $result) {
            $script:State.AzAvailable = $true
            Set-Status 'Ready. Sign in to begin.'
            return
        }
        $script:State.AzAvailable = [bool]$result.AzAvailable
        if (-not $result.AzAvailable) { Show-AzMissingBanner; Set-Status 'Az.Accounts not installed.'; return }
        if ($result.Account) {
            Enter-SignedInState -Account $result.Account -TenantId $result.TenantId -Context $result.Context
        }
        else {
            Set-Status 'Ready. Sign in to begin.'
        }
    }
}

function Export-VisibleCsv {
    $view = $script:State.View
    if ($null -eq $view) { return }
    $rows = @($view | ForEach-Object { $_.Source })
    $dialog = [Microsoft.Win32.SaveFileDialog]::new()
    $dialog.Title = 'Export PIM assignments'
    $dialog.Filter = 'CSV files (*.csv)|*.csv'
    $dialog.FileName = 'PIM-Assignment-Inventory-{0:yyyyMMdd-HHmm}.csv' -f (Get-Date)
    if ($dialog.ShowDialog($script:Window)) {
        Export-PimInventoryCsv -Row $rows -Path $dialog.FileName
        Set-Status "Exported $($rows.Count) row(s) to $($dialog.FileName)"
    }
}

function Open-HtmlReport {
    $view = $script:State.View
    if ($null -eq $view) { return }
    $rows = @($view | ForEach-Object { $_.Source })
    $dir = Join-Path $script:DataDir 'Reports'
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $path = Join-Path $dir ('PIM-Assignment-Inventory-{0:yyyyMMdd-HHmmss}.html' -f (Get-Date))
    Export-PimInventoryHtml -Inventory $script:State.Inventory -Row $rows -Path $path
    Start-Process -FilePath $path
    Set-Status "HTML report opened: $path"
}

#endregion

#region Startup ---------------------------------------------------------------

function Initialize-MainWindow {
    Import-Module $script:ModuleManifest -Force

    $loaded = Import-AppXaml -Path $script:XamlPath
    $script:Window = $loaded.Window
    $script:ui = $loaded.Controls
    $ui = $script:ui
    $script:WelcomeBodyText = $ui.WelcomeBody.Text

    $script:State = @{
        Mode = 'SignedOut'; Account = $null; TenantId = $null; Context = $null
        AzAvailable = $true; Busy = $false; RefreshQueued = $false
        Inventory = $null; GridItems = $null; View = $null
        LastRefresh = $null; NextRefresh = $null; SignInTask = $null
    }
    $script:SuppressAutoRefresh = $false
    $script:Progress = $null
    $script:BannerAction = $null
    Update-FilterState

    $script:Window.Title = "Azure PIM Inventory $script:AppVersion"
    if (Test-Path -LiteralPath $script:IconPath) {
        $script:Window.Icon = Get-IconImage -Path $script:IconPath -PreferredSize 32
        $big = Get-IconImage -Path $script:IconPath -PreferredSize 256
        $ui.HeaderIcon.Source = $big
        $ui.WelcomeIcon.Source = $big
    }

    # Restore settings.
    $script:Settings = Read-AppSetting
    $ui.TenantBox.Text = $script:Settings.Tenant
    Select-ComboTag $ui.WarningDaysCombo ([string]$script:Settings.WarningDays)
    $ui.IncludeFutureCheck.IsChecked = $script:Settings.IncludeFuture
    $ui.IncludeJitCheck.IsChecked    = $script:Settings.IncludeJit
    Select-ComboTag $ui.RefreshScheduleCombo $script:Settings.RefreshSchedule
    $ui.DailyTimeBox.Text = $script:Settings.DailyTime
    $ui.TenantPlaceholder.Visibility = if ($ui.TenantBox.Text) { 'Collapsed' } else { 'Visible' }

    # Header and sign-in.
    $ui.TenantBox.add_TextChanged({ $script:ui.TenantPlaceholder.Visibility = if ($script:ui.TenantBox.Text) { 'Collapsed' } else { 'Visible' } })
    $ui.TenantBox.add_KeyDown({ param($s, $e) if ($e.Key -eq 'Return') { Invoke-UiAction { Start-SignIn } } })
    $ui.SignInButton.add_Click({ Invoke-UiAction { Start-SignIn } })
    $ui.WelcomeSignInButton.add_Click({ Invoke-UiAction { Start-SignIn } })
    $ui.SignOutButton.add_Click({ Invoke-UiAction { Start-SignOut } })
    $ui.DemoButton.add_Click({ Invoke-UiAction { Start-Demo } })
    $ui.BannerActionButton.add_Click({ Invoke-UiAction { if ($script:BannerAction) { & $script:BannerAction } } })

    # Collection options: re-collect when they change.
    $ui.RefreshButton.add_Click({ Invoke-UiAction { Start-InventoryRefresh -Reason 'Manual' } })
    $ui.ManagementGroupCombo.add_SelectionChanged({
            Invoke-UiAction {
                if (-not $script:SuppressAutoRefresh -and $script:State.Mode -eq 'SignedIn' -and $null -ne $script:ui.ManagementGroupCombo.SelectedItem) {
                    Save-CurrentSetting
                    Start-InventoryRefresh -Reason 'Management group changed'
                }
            }
        })
    $ui.ManagementGroupCombo.add_KeyDown({ param($s, $e) if ($e.Key -eq 'Return') { Invoke-UiAction { Start-InventoryRefresh -Reason 'Manual' } } })
    $optionChanged = { Invoke-UiAction { Save-CurrentSetting; if ($script:State.Inventory) { Start-InventoryRefresh -Reason 'Options changed' } else { Update-Summary } } }
    $ui.WarningDaysCombo.add_SelectionChanged($optionChanged)
    $ui.IncludeFutureCheck.add_Click($optionChanged)
    $ui.IncludeJitCheck.add_Click($optionChanged)

    # Refresh schedule.
    $ui.RefreshScheduleCombo.add_SelectionChanged({ Invoke-UiAction { Update-Schedule; Save-CurrentSetting } })
    $ui.DailyTimeBox.add_TextChanged({ Invoke-UiAction { Update-Schedule } })
    $ui.DailyTimeBox.add_LostFocus({ Invoke-UiAction { Save-CurrentSetting } })

    # Filters (search is debounced so typing stays smooth on large estates).
    $script:SearchDebounce = [System.Windows.Threading.DispatcherTimer]::new()
    $script:SearchDebounce.Interval = [TimeSpan]::FromMilliseconds(250)
    $script:SearchDebounce.add_Tick({ $script:SearchDebounce.Stop(); Invoke-UiAction { Update-GridFilter } })
    $ui.SearchBox.add_TextChanged({
            $script:ui.SearchPlaceholder.Visibility = if ($script:ui.SearchBox.Text) { 'Collapsed' } else { 'Visible' }
            $script:SearchDebounce.Stop(); $script:SearchDebounce.Start()
        })
    $ui.StatusFilterCombo.add_SelectionChanged({ Invoke-UiAction { Update-GridFilter } })
    $ui.KindFilterCombo.add_SelectionChanged({ Invoke-UiAction { Update-GridFilter } })
    $ui.PrivilegedOnlyCheck.add_Click({ Invoke-UiAction { Update-GridFilter } })
    $ui.ClearFiltersButton.add_Click({
            Invoke-UiAction {
                $script:ui.SearchBox.Text = ''
                Select-ComboTag $script:ui.StatusFilterCombo ''
                Select-ComboTag $script:ui.KindFilterCombo ''
                $script:ui.PrivilegedOnlyCheck.IsChecked = $false
                Update-GridFilter
            }
        })

    # KPI cards act as quick filters.
    $cardFilters = @{
        CardTotal     = @{ Status = ''; Kind = ''; Priv = $false }
        CardExpired   = @{ Status = 'Expired'; Kind = ''; Priv = $false }
        CardExpiring  = @{ Status = 'Expiring'; Kind = ''; Priv = $false }
        CardPermanent = @{ Status = 'Permanent'; Kind = ''; Priv = $false }
        CardPending   = @{ Status = 'Pending'; Kind = ''; Priv = $false }
        CardEligible  = @{ Status = ''; Kind = 'Eligible'; Priv = $false }
        CardActive    = @{ Status = ''; Kind = 'Active'; Priv = $false }
        CardStanding  = @{ Status = ''; Kind = 'Active'; Priv = $true }
    }
    $script:CardFilters = $cardFilters
    foreach ($name in $cardFilters.Keys) {
        $ui[$name].Tag = $name
        $ui[$name].add_MouseLeftButtonUp({
                param($card, $e)
                Invoke-UiAction {
                    $f = $script:CardFilters[[string]$card.Tag]
                    Select-ComboTag $script:ui.StatusFilterCombo $f.Status
                    Select-ComboTag $script:ui.KindFilterCombo $f.Kind
                    $script:ui.PrivilegedOnlyCheck.IsChecked = $f.Priv
                    Update-GridFilter
                }
            })
    }

    # Exports and grid context menu.
    $ui.ExportCsvButton.add_Click({ Invoke-UiAction { Export-VisibleCsv } })
    $ui.HtmlReportButton.add_Click({ Invoke-UiAction { Open-HtmlReport } })
    $ui.CopyRowMenu.add_Click({ Invoke-UiAction { [System.Windows.Input.ApplicationCommands]::Copy.Execute($null, $script:ui.AssignmentGrid) } })
    $ui.CopyScopeMenu.add_Click({ Invoke-UiAction { $r = $script:ui.AssignmentGrid.SelectedItem; if ($r) { [System.Windows.Clipboard]::SetText($r.Scope) } } })
    $ui.CopyPrincipalIdMenu.add_Click({ Invoke-UiAction { $r = $script:ui.AssignmentGrid.SelectedItem; if ($r -and $r.PrincipalId) { [System.Windows.Clipboard]::SetText($r.PrincipalId) } } })
    $ui.OpenPortalMenu.add_Click({
            Invoke-UiAction {
                $r = $script:ui.AssignmentGrid.SelectedItem
                if ($r) {
                    $url = Get-PortalScopeUrl -Scope $r.Scope -TenantId ([string]$script:State.TenantId)
                    if ($url) { Start-Process -FilePath $url }
                }
            }
        })

    # Timers: one polls background work, one drives the refresh schedule.
    $script:PollTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:PollTimer.Interval = [TimeSpan]::FromMilliseconds(250)
    $script:PollTimer.add_Tick({
            Invoke-UiAction {
                Receive-AppTask
                $p = $script:Progress
                if ($null -ne $p -and $script:State.Busy) {
                    $percent = [int]$p['Percent']
                    $script:ui.BusyBar.IsIndeterminate = ($percent -lt 0)
                    if ($percent -ge 0) { $script:ui.BusyBar.Value = $percent }
                    if ($p['Message']) { $script:ui.StatusText.Text = [string]$p['Message'] }
                }
            }
        })

    $script:ScheduleTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:ScheduleTimer.Interval = [TimeSpan]::FromSeconds(15)
    $script:ScheduleTimer.add_Tick({
            Invoke-UiAction {
                $next = $script:State.NextRefresh
                if ($null -ne $next -and (Get-Date) -ge $next -and $script:State.Mode -ne 'SignedOut' -and -not $script:State.Busy) {
                    $script:State.NextRefresh = $null
                    Start-InventoryRefresh -Reason 'Scheduled'
                }
                Update-ScheduleText
            }
        })

    $script:Window.Dispatcher.add_UnhandledException({
            param($dispatcher, $e)
            Write-AppLog "UNHANDLED: $($e.Exception)"
            $e.Handled = $true
            try { Show-AppError -Message $e.Exception.Message } catch { $null = $_ }
        })

    $script:Window.add_Closing({
            Save-CurrentSetting
            $script:PollTimer.Stop()
            $script:ScheduleTimer.Stop()
            foreach ($t in @($script:Tasks)) { Stop-AppTask -Task $t }
        })

    Update-HeaderState
    Update-Schedule
    return $script:Window
}

function Start-PimDashboard {
    param([switch]$Demo)

    if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
        throw 'The dashboard must run on an STA thread. Start it with Start-PimDashboard.cmd, or "powershell -STA".'
    }

    $window = Initialize-MainWindow
    $script:PollTimer.Start()
    $script:ScheduleTimer.Start()

    $script:StartInDemo = [bool]$Demo
    $window.add_ContentRendered({
            Invoke-UiAction { if ($script:StartInDemo) { Start-Demo } else { Start-SessionCheck } }
        })

    Write-AppLog "App $script:AppVersion started on PowerShell $($PSVersionTable.PSVersion)."
    [void]$window.ShowDialog()
    Write-AppLog 'App closed.'
}

#endregion

# Run only when executed, not when dot-sourced (tests load the functions).
if ($MyInvocation.InvocationName -ne '.') {
    try {
        Start-PimDashboard -Demo:$Demo
    }
    catch {
        Write-AppLog "FATAL: $($_.Exception)"
        try {
            Add-Type -AssemblyName PresentationFramework
            [void][System.Windows.MessageBox]::Show("Azure PIM Inventory could not start:`n`n$($_.Exception.Message)`n`nDetails: $script:LogPath",
                'Azure PIM Inventory', 'OK', 'Error')
        }
        catch { Write-Error $_ }
        exit 1
    }
}
