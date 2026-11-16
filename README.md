# Azure PIM Assignment Inventory

Exports every PIM role assignment (eligible and active) beneath an Azure management
group to CSV and a self-contained HTML dashboard.

Built for PIRSA, September 2026.

## Why this exists

The Azure portal cannot show you every PIM assignment across a management group
hierarchy in one view. You have to inspect each scope individually, which means
time-bound assignments quietly reach expiry unnoticed. This produces the single
consolidated list, with expiry dates, that the portal does not.

## Quick start

```powershell
Install-Module Az.Accounts -Scope CurrentUser   # if not already present
Connect-AzAccount
.\Export-AzPimAssignmentInventory.ps1 -RootManagementGroupId 'your-mg-id'
```

Use your actual management-group ID, not the example `mg-pirsa-root`.
Hierarchies containing only the root management group and no child subscriptions
are supported.

Two files are written to the current directory:

- `PIM-Assignment-Inventory-<timestamp>.csv` — for Excel, Power BI or audit evidence
- `PIM-Assignment-Inventory-<timestamp>.html` — dashboard with filtering and sorting

Preview the output format without connecting to Azure:

```powershell
.\Export-AzPimAssignmentInventory.ps1 -OfflineSampleData .\sample-data.json
```

## Requirements

| | |
|---|---|
| PowerShell | 5.1 or 7.x |
| Module | `Az.Accounts` only. `Az.Resources` is **not** required. |
| Azure rights | **Reader** at the root management group is normally enough, because listing needs `Microsoft.Authorization/*/read`. Microsoft documents Owner or User Access Administrator for the PIM ARM API generally, so if calls return 403, grant User Access Administrator. |
| Entra rights | **None.** No Graph permissions, no Global Administrator, no Privileged Role Administrator. |
| Licence | PIM requires Microsoft Entra ID P2 or Microsoft Entra ID Governance. |

## Parameters

| Parameter | Default | Purpose |
|---|---|---|
| `-RootManagementGroupId` | required | Management group ID (the name, not the display name). |
| `-OutputPath` | current directory | Where to write the CSV and HTML. |
| `-ExpiryWarningDays` | 30 | Assignments expiring within this many days are flagged `Expiring`. |
| `-IncludeActivations` | off | Include transient just-in-time activations. |
| `-IncludeAncestorScopes` | off | Include assignments inherited from above the root management group. |
| `-IncludeFutureAssignments` | off | Also include assignments scheduled to *start* in the future (queries the `*Schedules` resources instead of `*ScheduleInstances`). |
| `-SkipHtml` | off | Produce the CSV only. |
| `-OfflineSampleData` | — | Preview the output format using a sample JSON file. |

## Output columns

`PrincipalName`, `PrincipalType`, `PrincipalEmail`, `PrincipalId`, `RoleName`,
`RoleType`, `AssignmentKind`, `MemberType`, `ScopeName`, `ScopeType`, `Scope`,
`StartDateUtc`, `ExpiryDateUtc`, `DaysRemaining`, `Status`, `Condition`, `InstanceId`

`Status` is one of `Expired`, `Expiring`, `Active` or `Permanent`.
`AssignmentKind` is one of `Eligible`, `Active` or `Activation (JIT)`.

## Things worth knowing

**It uses Azure Resource Manager only.** PIM for Azure resources is an ARM API, not
a Microsoft Graph API. Principal, role and scope display names come from the ARM
`expandedProperties` payload, so no directory permissions are needed.

**Do not try to rebuild this on Azure Resource Graph.** Resource Graph cannot see
eligible assignments — its `authorizationresources` table has no eligible-schedule
type. A report built on it returns permanent assignments only and omits every
eligible assignment silently, with no error.

**Coverage is explicit rather than assumed.** ARM returns assignments at, above and
below a queried scope, but Microsoft documents that behaviour from a subscription
scope, not from a management group root. The script therefore enumerates the root
management group plus all descendant management groups and subscriptions, queries
each, and de-duplicates on instance ID. Slower, but correct either way.

**Just-in-time activations are excluded by default.** An activated eligible
assignment appears as an active instance with `assignmentType = 'Activated'` and
would otherwise add rows that disappear within hours.

**Assignments made outside PIM are included.** The active-assignment API also
returns permanent RBAC granted outside PIM entirely, so standing access that
bypassed PIM shows up in the report. This is usually the most interesting part.

**Future-dated assignments are hidden by default.** Microsoft: *"schedule instances
only include assignments that are active at the current time, schedules also include
assignments that will become active in the future."* The script queries schedule
instances by default, so it reports the estate as it stands today. Add
`-IncludeFutureAssignments` to see assignments scheduled to start later.

**Groups are not expanded.** Where a group holds the assignment, the group is
reported as the principal. Expanding to individual members would require Microsoft
Graph directory permissions.

**Partial failures are visible.** If a scope cannot be read the run continues, and
the affected scopes are listed in a warning banner at the top of the HTML report,
so an incomplete inventory can never be mistaken for a clean one.

## Scheduling it

For unattended runs, deploy to an Azure Automation account with a system-assigned
managed identity granted Reader at the root management group, import `Az.Accounts`,
and replace `Connect-AzAccount` with `Connect-AzAccount -Identity`. Write the output
to a storage account or SharePoint library, and point Power BI at that folder if a
refreshable dashboard is wanted.

## Files

| File | Purpose |
|---|---|
| `Export-AzPimAssignmentInventory.ps1` | The script. |
| `sample-data.json` | Sample ARM payloads for offline preview. Not real data. |
| `SampleOutput\` | Example CSV and HTML generated from the sample data. |
