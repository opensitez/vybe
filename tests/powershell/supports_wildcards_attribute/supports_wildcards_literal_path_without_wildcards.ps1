# vybe-test: powershell/supports_wildcards_attribute/supports_wildcards_literal_path_without_wildcards
# A pattern passed to a SupportsWildcards parameter with no wildcard characters performs exact literal matching
function MatchExactLiteral {
    [CmdletBinding()]
    param(
        [SupportsWildcards()]
        [string]$Pattern
    )
    process {
        $pool = @("config.json", "config.json.bak", "settings.json")
        $wp = [System.Management.Automation.WildcardPattern]::new($Pattern)
        $matches = @($pool | Where-Object { $wp.IsMatch($_) })
        return $matches -join ", "
    }
}

$matched = MatchExactLiteral -Pattern "config.json"

if ($matched -ne "config.json") {
    Write-Host "FAIL: exact literal match failed: '$matched'"
    exit 1
}

Write-Host "PASS"
exit 0
