# vybe-test: powershell/parameters_alias_attribute/alias_resolves_in_get_help_or_parameter_set
function Test-AliasLookup {
    [CmdletBinding()]
    param([Alias("Target")][string]$Destination)
    return $Destination
}
$cmd = Get-Command Test-AliasLookup
$param = $cmd.Parameters["Destination"]
if ($param.Aliases -notcontains "Target") {
    Write-Host "FAIL: Aliased parameter lookup in Parameters dictionary failed"
    exit 1
}
Write-Host "PASS"
exit 0
