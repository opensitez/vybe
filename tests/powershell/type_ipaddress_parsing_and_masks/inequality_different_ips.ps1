# vybe-test: powershell/type_ipaddress_parsing_and_masks/inequality_different_ips
$ip1 = [System.Net.IPAddress]::Parse("10.0.0.1")
$ip2 = [System.Net.IPAddress]::Parse("10.0.0.2")
if ($ip1 -eq $ip2) {
    Write-Host "FAIL: different IP addresses must compare unequal"
    exit 1
}
Write-Host "PASS"
exit 0
