# vybe-test: powershell/supports_wildcards_attribute/supports_wildcards_case_insensitive_matching
# WildcardPattern default matching is case-insensitive for strings received via SupportsWildcards parameter
function MatchCaseInsensitive {
    [CmdletBinding()]
    param(
        [SupportsWildcards()]
        [string]$Pattern
    )
    process {
        $pool = @("README.MD", "license.txt", "CONTRIBUTING.MD")
        $wp = [System.Management.Automation.WildcardPattern]::new($Pattern, [System.Management.Automation.WildcardOptions]::IgnoreCase)
        $matches = @($pool | Where-Object { $wp.IsMatch($_) })
        return $matches -join ", "
    }
}

$matched = MatchCaseInsensitive -Pattern "*.md"

if ($matched -ne "README.MD, CONTRIBUTING.MD") {
    Write-Host "FAIL: case-insensitive wildcard match failed: '$matched'"
    exit 1
}

Write-Host "PASS"
exit 0
