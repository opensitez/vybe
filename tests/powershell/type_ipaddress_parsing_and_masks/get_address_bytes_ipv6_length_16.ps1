# vybe-test: powershell/type_ipaddress_parsing_and_masks/get_address_bytes_ipv6_length_16
$ip = [System.Net.IPAddress]::Parse("::1")
$bytes = $ip.GetAddressBytes()
if ($bytes.Length -ne 16 -or $bytes[15] -ne 1) {
    Write-Host "FAIL: IPv6 GetAddressBytes mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
