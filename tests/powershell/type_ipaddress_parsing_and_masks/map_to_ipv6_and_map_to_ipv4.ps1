# vybe-test: powershell/type_ipaddress_parsing_and_masks/map_to_ipv6_and_map_to_ipv4
$v4 = [System.Net.IPAddress]::Parse("192.0.2.1")
$v6 = $v4.MapToIPv6()
$backToV4 = $v6.MapToIPv4()
if ($backToV4.ToString() -ne "192.0.2.1") {
    Write-Host "FAIL: MapToIPv6 / MapToIPv4 roundtrip failed"
    exit 1
}
Write-Host "PASS"
exit 0
