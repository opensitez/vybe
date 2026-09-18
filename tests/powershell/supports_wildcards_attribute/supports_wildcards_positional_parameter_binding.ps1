# vybe-test: powershell/supports_wildcards_attribute/supports_wildcards_positional_parameter_binding
# A positional parameter decorated with [SupportsWildcards()] accepts positional arguments cleanly
function MatchPositionalPattern {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)]
        [SupportsWildcards()]
        [string]$Pattern
    )
    process {
        return "PositionalPattern:$Pattern"
    }
}

$boundVal = MatchPositionalPattern "asset-*.png"

if ($boundVal -ne "PositionalPattern:asset-*.png") {
    Write-Host "FAIL: positional wildcard parameter binding failed: '$boundVal'"
    exit 1
}

Write-Host "PASS"
exit 0
