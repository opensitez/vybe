# vybe-test: powershell/type_ipaddress_parsing_and_masks/is_ipv6_multicast
$ip1 = [System.Net.IPAddress]::Parse("ff02::1")
$ip2 = [System.Net.IPAddress]::Parse("fe80::1")
if (-not $ip1.IsIPv6Multicast -or $ip2.IsIPv6Multicast) {
    Write-Host "FAIL: IsIPv6Multicast check failed"
    exit 1
}
Write-Host "PASS"
exit 0
