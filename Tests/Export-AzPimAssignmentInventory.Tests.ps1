$inventoryScript = Join-Path (Split-Path -Parent $PSScriptRoot) 'Export-AzPimAssignmentInventory.ps1'
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile(
    $inventoryScript, [ref]$tokens, [ref]$parseErrors
)
if ($parseErrors.Count -gt 0) {
    throw ($parseErrors -join [Environment]::NewLine)
}

$functionAst = $ast.Find({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
    $node.Name -eq 'Get-TargetScope'
}, $true)
. ([scriptblock]::Create($functionAst.Extent.Text))

function Invoke-ArmList {
    param([string]$Path)
    throw 'Azure requests must be mocked.'
}

Describe 'Azure PIM inventory hierarchy collection' {
    BeforeEach {
        $script:MgApiVersion = '2020-05-01'
    }

    It 'keeps the root as a collection when no descendants are returned' {
        Set-StrictMode -Version Latest
        Mock Invoke-ArmList { return , @() }

        $scopes = Get-TargetScope -ManagementGroupId 'mg-test'

        $scopes.Count | Should Be 1
        $scopes[0] | Should Be '/providers/Microsoft.Management/managementGroups/mg-test'
        Assert-MockCalled Invoke-ArmList -Times 1 -Exactly -Scope It -ParameterFilter {
            $Path -eq '/providers/Microsoft.Management/managementGroups/mg-test/descendants?api-version=2020-05-01'
        }
    }

    It 'preserves descendant traversal and removes duplicate scopes' {
        Set-StrictMode -Version Latest
        Mock Invoke-ArmList {
            return , @(
                [pscustomobject]@{ type = 'Microsoft.Management/managementGroups'; name = 'mg-child' }
                [pscustomobject]@{ type = 'Microsoft.Management/managementGroups/subscriptions'; name = '00000000-0000-0000-0000-000000000003' }
                [pscustomobject]@{ type = 'Microsoft.Management/managementGroups/subscriptions'; name = '00000000-0000-0000-0000-000000000003' }
            )
        }

        $scopes = Get-TargetScope -ManagementGroupId 'mg-test'

        $scopes.Count | Should Be 3
        $scopes[0] | Should Be '/providers/Microsoft.Management/managementGroups/mg-test'
        $scopes[1] | Should Be '/providers/Microsoft.Management/managementGroups/mg-child'
        $scopes[2] | Should Be '/subscriptions/00000000-0000-0000-0000-000000000003'
    }
}
