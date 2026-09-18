# vybe-test: powershell/supports_wildcards_attribute/supports_wildcards_multiple_wildcard_parameters_in_function
# An advanced function can declare multiple parameters each decorated with [SupportsWildcards()]
function MatchMultipleWildcardArgs {
    [CmdletBinding()]
    param(
        [SupportsWildcards()]
        [string]$IncludePattern,

        [SupportsWildcards()]
        [string]$ExcludePattern
    )
    process {
        $cmd = Get-Command MatchMultipleWildcardArgs
        $incParam = $cmd.Parameters["IncludePattern"]
        $excParam = $cmd.Parameters["ExcludePattern"]

        $incHas = ($incParam.Attributes | Where-Object { $_ -is [System.Management.Automation.SupportsWildcardsAttribute] }) -ne $null
        $excHas = ($excParam.Attributes | Where-Object { $_ -is [System.Management.Automation.SupportsWildcardsAttribute] }) -ne $null

        return "Inc:$incHas,Exc:$excHas,IncVal:$IncludePattern,ExcVal:$ExcludePattern"
    }
}

$summary = MatchMultipleWildcardArgs -IncludePattern "*.cs" -ExcludePattern "*Test*"

if ($summary -ne "Inc:True,Exc:True,IncVal:*.cs,ExcVal:*Test*") {
    Write-Host "FAIL: multiple wildcard parameter summary mismatch: '$summary'"
    exit 1
}

Write-Host "PASS"
exit 0
