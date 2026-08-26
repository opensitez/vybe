# vybe-test: powershell/type_ipaddress_parsing_and_masks/get_address_bytes_ipv4_length_4
$ip = [System.Net.IPAddress]::Parse("192.168.0.10")
$bytes = $ip.GetAddressBytes()
if ($bytes.Length -ne 4 -or $bytes[0] -ne 192 -or $bytes[1] -ne 168 -or $bytes[2] -ne 0 -or $bytes[3] -ne 10) {
    Write-Host "FAIL: IPv4 GetAddressBytes mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
