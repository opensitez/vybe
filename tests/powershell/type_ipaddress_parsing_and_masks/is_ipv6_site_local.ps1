# vybe-test: powershell/type_ipaddress_parsing_and_masks/is_ipv6_site_local
$ip1 = [System.Net.IPAddress]::Parse("fec0::1")
$ip2 = [System.Net.IPAddress]::Parse("2001:db8::1")
if (-not $ip1.IsIPv6SiteLocal -or $ip2.IsIPv6SiteLocal) {
    Write-Host "FAIL: IsIPv6SiteLocal check failed"
    exit 1
}
Write-Host "PASS"
exit 0
