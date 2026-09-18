# vybe-test: powershell/supports_wildcards_attribute/supports_wildcards_empty_pattern_matching
# An empty string pattern passed to a SupportsWildcards parameter matches only empty strings
function MatchEmptyPattern {
    [CmdletBinding()]
    param(
        [SupportsWildcards()]
        [string]$Pattern
    )
    process {
        $wp = [System.Management.Automation.WildcardPattern]::new($Pattern)
        $matchEmpty = $wp.IsMatch("")
        $matchNonEmpty = $wp.IsMatch("data")
        return "MatchEmpty:$matchEmpty,MatchNonEmpty:$matchNonEmpty"
    }
}

$res = MatchEmptyPattern -Pattern ""

if ($res -ne "MatchEmpty:True,MatchNonEmpty:False") {
    Write-Host "FAIL: empty pattern evaluation mismatch: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
