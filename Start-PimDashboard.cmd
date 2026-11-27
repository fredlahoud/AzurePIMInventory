@echo off
rem Azure PIM Inventory - double-click to open the desktop dashboard.
rem Uses PowerShell 7 when installed (recommended), otherwise Windows PowerShell 5.1.
setlocal
set "APP=%~dp0App\PimDashboard.ps1"

where pwsh.exe >nul 2>&1
if %ERRORLEVEL%==0 (
    start "" pwsh.exe -NoProfile -NoLogo -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "%APP%" %*
) else (
    start "" powershell.exe -NoProfile -NoLogo -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "%APP%" %*
)
endlocal
