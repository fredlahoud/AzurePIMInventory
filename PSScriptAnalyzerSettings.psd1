@{
    # Errors and warnings fail CI. Information-level findings are advisory only.
    Severity     = @('Error', 'Warning')

    ExcludeRules = @(
        # The script is an interactive/console tool; Write-Host is deliberate for
        # coloured progress output and is captured by the Information stream in
        # PowerShell 5+ (including Azure Automation).
        'PSAvoidUsingWriteHost'
    )
}
