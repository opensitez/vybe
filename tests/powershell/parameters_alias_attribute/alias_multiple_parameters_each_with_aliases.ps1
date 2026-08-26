# vybe-test: powershell/parameters_alias_attribute/alias_multiple_parameters_each_with_aliases
function Connect-EndPoint {
    param(
        [Alias("H")][string]$HostName,
        [Alias("P")][int]$Port
    )
    return "$HostName`:$Port"
}
$res = Connect-EndPoint -H "localhost" -P 443
if ($res -ne "localhost:443") {
    Write-Host "FAIL: Multiple parameters with aliases failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
