## Summary

<!-- What changes, and why. -->

## Testing

- [ ] `Invoke-Pester ./Tests` passes locally (Pester 5)
- [ ] `Invoke-ScriptAnalyzer -Path . -Recurse -Settings ./PSScriptAnalyzerSettings.psd1` is clean
- [ ] Offline preview renders: `./Export-AzPimAssignmentInventory.ps1 -OfflineSampleData ./sample-data.json`
- [ ] Tested against a live management group (if ARM calls changed)

## Output contract

- [ ] No change to CSV columns, or the change is noted in `CHANGELOG.md` and `README.md`
