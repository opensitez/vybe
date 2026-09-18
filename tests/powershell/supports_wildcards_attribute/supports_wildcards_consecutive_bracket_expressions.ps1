# vybe-test: powershell/supports_wildcards_attribute/supports_wildcards_consecutive_bracket_expressions
# Consecutive bracket expressions [a-z][0-9] match multi-character combinations sequentially
function MatchCombinedBrackets {
    [CmdletBinding()]
    param(
        [SupportsWildcards()]
        [string]$Pattern
    )
    process {
        $items = @("item-a1", "item-b2", "item-99", "item-ax")
        $wp = [System.Management.Automation.WildcardPattern]::new($Pattern)
        $matches = @($items | Where-Object { $wp.IsMatch($_) })
        return $matches -join ", "
    }
}

$result = MatchCombinedBrackets -Pattern "item-[a-z][0-9]"

if ($result -ne "item-a1, item-b2") {
    Write-Host "FAIL: consecutive bracket expressions matching failed: '$result'"
    exit 1
}

Write-Host "PASS"
exit 0
