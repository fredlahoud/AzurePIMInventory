#Requires -Version 5.1
<#
.SYNOPSIS
    Adds "Azure PIM Inventory" shortcuts to the Desktop and Start menu.

.DESCRIPTION
    The shortcuts start the dashboard directly with a hidden PowerShell host, so no
    console window appears. Re-run after moving the repository folder.
    Per-user only: no administrator rights are needed.

.PARAMETER NoDesktop
    Skip the Desktop shortcut.

.PARAMETER NoStartMenu
    Skip the Start menu shortcut.

.PARAMETER Remove
    Remove the shortcuts instead of creating them.

.EXAMPLE
    .\Install-DesktopShortcut.ps1
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [switch]$NoDesktop,
    [switch]$NoStartMenu,
    [switch]$Remove
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$name    = 'Azure PIM Inventory'
$appPath = Join-Path $PSScriptRoot 'App\PimDashboard.ps1'
$icon    = Join-Path $PSScriptRoot 'App\assets\app.ico'

$targets = @()
if (-not $NoDesktop)   { $targets += Join-Path ([Environment]::GetFolderPath('Desktop')) "$name.lnk" }
if (-not $NoStartMenu) { $targets += Join-Path ([Environment]::GetFolderPath('Programs')) "$name.lnk" }

if ($Remove) {
    foreach ($t in $targets) {
        if ((Test-Path -LiteralPath $t) -and $PSCmdlet.ShouldProcess($t, 'Remove shortcut')) {
            Remove-Item -LiteralPath $t
            Write-Host "Removed $t"
        }
    }
    return
}

if (-not (Test-Path -LiteralPath $appPath)) { throw "App not found at $appPath" }

# Prefer PowerShell 7; fall back to Windows PowerShell 5.1, which every Windows has.
$hostExe = (Get-Command pwsh.exe -ErrorAction SilentlyContinue | Select-Object -First 1).Source
if (-not $hostExe) { $hostExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe' }

$shell = New-Object -ComObject WScript.Shell
foreach ($t in $targets) {
    if (-not $PSCmdlet.ShouldProcess($t, 'Create shortcut')) { continue }
    $lnk = $shell.CreateShortcut($t)
    $lnk.TargetPath       = $hostExe
    $lnk.Arguments        = "-NoProfile -NoLogo -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$appPath`""
    $lnk.WorkingDirectory = $PSScriptRoot
    $lnk.IconLocation     = "$icon,0"
    $lnk.Description      = 'Inventory of Azure PIM eligible and active role assignments'
    $lnk.WindowStyle      = 7   # minimised: avoids a console flash before -WindowStyle Hidden applies
    $lnk.Save()
    Write-Host "Created $t"
}
