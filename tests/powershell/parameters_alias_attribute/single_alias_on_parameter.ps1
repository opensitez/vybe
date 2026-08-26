# vybe-test: powershell/parameters_alias_attribute/single_alias_on_parameter
function Get-ServerInfo {
    param(
        [Alias("ComputerName")]
        [string]$Server
    )
    return "Server:$Server"
}
$res = Get-ServerInfo -ComputerName "db01"
if ($res -ne "Server:db01") {
    Write-Host "FAIL: Single alias binding failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
