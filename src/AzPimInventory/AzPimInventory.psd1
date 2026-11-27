@{
    RootModule           = 'AzPimInventory.psm1'
    ModuleVersion        = '2.0.0'
    GUID                 = 'e6ca2451-6ea6-4a5d-95db-0b9cacde9f03'
    Author               = 'fredlahoud'
    Description          = 'Inventories Azure PIM role assignments (eligible and active) beneath a management group using Azure Resource Manager only. No Microsoft Graph permissions required.'
    PowerShellVersion    = '5.1'
    CompatiblePSEditions = @('Desktop', 'Core')

    # Az.Accounts is needed for live collection only; offline mode and the export
    # functions work without it, so it is checked at run time instead.
    FunctionsToExport    = @(
        'Get-PimInventory'
        'Get-PimManagementGroup'
        'Export-PimInventoryCsv'
        'ConvertTo-PimInventoryHtml'
        'Export-PimInventoryHtml'
    )
    CmdletsToExport      = @()
    VariablesToExport    = @()
    AliasesToExport      = @()

    PrivateData          = @{
        PSData = @{
            Tags       = @('Azure', 'PIM', 'RBAC', 'Governance', 'Security', 'Inventory')
            ProjectUri = 'https://github.com/fredlahoud/AzurePIMInventory'
        }
    }
}
