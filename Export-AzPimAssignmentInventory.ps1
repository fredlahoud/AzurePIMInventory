#Requires -Version 5.1
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
      will become active in the future". Such rows get the status 'Pending'.

    * Standing privileged access is called out. Active (non-JIT) assignments of a
      role listed in -PrivilegedRoleName are counted on the dashboard, because
      standing Owner / User Access Administrator access is the finding most
      security reviews are looking for.

.PARAMETER RootManagementGroupId
    Management group ID (the name, not the display name) to report from.

.PARAMETER OfflineSampleData
    Path to a JSON file of sample instances, used to preview the output format
    without connecting to Azure. Intended for validation and demonstration.

.PARAMETER ReferenceDateUtc
    Offline mode only. The instant, in UTC, against which expiry is evaluated.
    Defaults to now. Lets sample output and tests be reproduced exactly.

.PARAMETER OutputPath
    Directory for the CSV and HTML output. Defaults to the current directory.

.PARAMETER ExpiryWarningDays
    Assignments expiring within this many days are flagged 'Expiring'. Default 30.

.PARAMETER PrivilegedRoleName
    Role display names treated as privileged. Defaults to the roles the Azure
    portal lists as privileged administrator roles: Owner, Contributor,
    User Access Administrator and Role Based Access Control Administrator.

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

.PARAMETER PassThru
    Also write the report rows to the pipeline, for further processing.

.PARAMETER FailOnScopeError
    Throw (non-zero exit) after the reports are written if any scope could not be
    read. Use in pipelines and Automation runbooks so an incomplete inventory fails
    the job rather than passing silently.

.EXAMPLE
    .\Export-AzPimAssignmentInventory.ps1 -RootManagementGroupId 'mg-pirsa-root'

.EXAMPLE
    .\Export-AzPimAssignmentInventory.ps1 -RootManagementGroupId 'mg-pirsa-root' `
        -ExpiryWarningDays 60 -OutputPath 'C:\Reports'

.EXAMPLE
    $rows = .\Export-AzPimAssignmentInventory.ps1 -RootManagementGroupId 'mg-pirsa-root' -PassThru -SkipHtml
    $rows | Where-Object { $_.PrivilegedRole -eq 'Yes' -and $_.AssignmentKind -eq 'Active' }

.EXAMPLE
    .\Export-AzPimAssignmentInventory.ps1 -OfflineSampleData .\sample-data.json -ReferenceDateUtc '2026-09-11 01:36'

.NOTES
    Requires : The AzPimInventory module in src\ (shipped with this script).
               Az.Accounts (Connect-AzAccount first). Az.Resources is not required.
               Offline mode needs no modules.
    Rights   : Reader at the root management group is normally sufficient, because
               listing needs Microsoft.Authorization/*/read. Microsoft documents
               Owner or User Access Administrator for the PIM ARM API generally, so
               if calls return 403, grant User Access Administrator instead.
    Licence  : PIM requires Microsoft Entra ID P2 or Microsoft Entra ID Governance.
#>

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

    [ValidateNotNullOrEmpty()]
    [string]$OutputPath = (Get-Location).Path,

    [ValidateRange(1, 365)]
    [int]$ExpiryWarningDays = 30,

    # Defaults to the module's list: Owner, Contributor, User Access Administrator,
    # Role Based Access Control Administrator.
    [ValidateNotNull()]
    [string[]]$PrivilegedRoleName,

    [switch]$IncludeActivations,

    [switch]$IncludeAncestorScopes,

    [switch]$IncludeFutureAssignments,

    [switch]$SkipHtml,

    [switch]$PassThru,

    [switch]$FailOnScopeError
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'src/AzPimInventory/AzPimInventory.psd1') -Force

$collect = @{
    ExpiryWarningDays        = $ExpiryWarningDays
    IncludeActivations       = $IncludeActivations
    IncludeAncestorScopes    = $IncludeAncestorScopes
    IncludeFutureAssignments = $IncludeFutureAssignments
}
if ($PSBoundParameters.ContainsKey('PrivilegedRoleName')) { $collect.PrivilegedRoleName = $PrivilegedRoleName }

if ($PSCmdlet.ParameterSetName -eq 'Offline') {
    $collect.OfflineSampleData = $OfflineSampleData
    if ($PSBoundParameters.ContainsKey('ReferenceDateUtc')) { $collect.ReferenceDateUtc = $ReferenceDateUtc }
}
else {
    $collect.RootManagementGroupId = $RootManagementGroupId
}

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

Write-Host ''
Write-Host 'Azure PIM assignment inventory' -ForegroundColor White

# Module-qualified so a same-named function in the caller's session is never used.
$inventory = AzPimInventory\Get-PimInventory @collect
$rows      = @($inventory.Rows)
$stamp     = $inventory.GeneratedUtc.ToString('yyyyMMdd-HHmm', [cultureinfo]::InvariantCulture)

if ($rows.Count -eq 0) {
    Write-Warning 'No PIM assignments were found. Verify the management group ID and that the signed-in identity can read role assignments at that scope.'
}

$csvPath = Join-Path $OutputPath "PIM-Assignment-Inventory-$stamp.csv"
AzPimInventory\Export-PimInventoryCsv -Row $rows -Path $csvPath
Write-Host "  CSV written to $csvPath" -ForegroundColor Cyan

if (-not $SkipHtml) {
    $htmlPath = Join-Path $OutputPath "PIM-Assignment-Inventory-$stamp.html"
    AzPimInventory\Export-PimInventoryHtml -Inventory $inventory -Path $htmlPath
    Write-Host "  HTML dashboard written to $htmlPath" -ForegroundColor Cyan
}

Write-Host ''
Write-Host 'Summary' -ForegroundColor White
$rows | Group-Object Status | Sort-Object Name | ForEach-Object {
    Write-Host ("  {0,-10} {1}" -f $_.Name, $_.Count)
}
$standingCount = @($rows | Where-Object { $_.AssignmentKind -eq 'Active' -and $_.PrivilegedRole -eq 'Yes' }).Count
if ($standingCount -gt 0) {
    Write-Host ("  {0} active assignment(s) of a privileged role (standing access)." -f $standingCount) -ForegroundColor Yellow
}

$scopeErrorCount = @($inventory.ScopeErrors).Count
if ($scopeErrorCount -gt 0) {
    Write-Host ''
    Write-Warning "$scopeErrorCount scope(s) could not be read. The inventory is incomplete - see the HTML report for detail."
}

Write-Host ''

if ($PassThru) { $rows }

if ($FailOnScopeError -and $scopeErrorCount -gt 0) {
    throw "$scopeErrorCount scope(s) could not be read; the inventory is incomplete. Reports were written to $OutputPath."
}
