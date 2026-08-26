# vybe-test: powershell/type_ipaddress_parsing_and_masks/is_ipv6_link_local
$ip1 = [System.Net.IPAddress]::Parse("fe80::1ff:fe23:4567:890a")
$ip2 = [System.Net.IPAddress]::Parse("2001:db8::1")
if (-not $ip1.IsIPv6LinkLocal -or $ip2.IsIPv6LinkLocal) {
    Write-Host "FAIL: IsIPv6LinkLocal check failed"
    exit 1
}
Write-Host "PASS"
exit 0
