# vybe-test: powershell/supports_wildcards_attribute/supports_wildcards_coexistence_with_alias_attribute
# SupportsWildcards and Alias attributes coexist without conflict on the same parameter
function ResolveWildcardViaAlias {
    [CmdletBinding()]
    param(
        [SupportsWildcards()]
        [Alias("Filter")]
        [string]$Pattern
    )
    process {
        return "BoundPattern:$Pattern"
    }
}

$res = ResolveWildcardViaAlias -Filter "service-*-active"

if ($res -ne "BoundPattern:service-*-active") {
    Write-Host "FAIL: alias binding with SupportsWildcards failed: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
