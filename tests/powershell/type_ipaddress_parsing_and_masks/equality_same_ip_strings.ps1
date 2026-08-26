# vybe-test: powershell/type_ipaddress_parsing_and_masks/equality_same_ip_strings
$ip1 = [System.Net.IPAddress]::Parse("10.10.10.10")
$ip2 = [System.Net.IPAddress]::Parse("10.10.10.10")
if ($ip1 -ne $ip2) {
    Write-Host "FAIL: identical IP addresses must compare equal"
    exit 1
}
Write-Host "PASS"
exit 0
