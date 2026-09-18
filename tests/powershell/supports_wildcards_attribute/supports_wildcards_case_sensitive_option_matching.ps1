# vybe-test: powershell/supports_wildcards_attribute/supports_wildcards_case_sensitive_option_matching
# WildcardOptions.None enables case-sensitive wildcard evaluation
function MatchCaseSensitive {
    [CmdletBinding()]
    param(
        [SupportsWildcards()]
        [string]$Pattern
    )
    process {
        $files = @("data.XML", "data.xml", "data.Xml")
        $wp = [System.Management.Automation.WildcardPattern]::new($Pattern, [System.Management.Automation.WildcardOptions]::None)
        $matches = @($files | Where-Object { $wp.IsMatch($_) })
        return $matches -join ", "
    }
}

$matched = MatchCaseSensitive -Pattern "*.XML"

if ($matched -ne "data.XML") {
    Write-Host "FAIL: case-sensitive wildcard match failed, expected 'data.XML', got: '$matched'"
    exit 1
}

Write-Host "PASS"
exit 0
