# vybe-test: powershell/dynamic_assembly_type_resolution/resolve_net_utility_accelerators
$tGuid = [type]"guid"
$tVer = [type]"version"
$tTs = [type]"timespan"
$tIp = [type]"ipaddress"
if ($tGuid -ne [guid] -or $tVer -ne [version] -or $tTs -ne [timespan] -or $tIp -ne [System.Net.IPAddress]) {
    Write-Host "FAIL: .NET utility accelerator resolution failed"
    exit 1
}
Write-Host "PASS"
exit 0
