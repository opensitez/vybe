# vybe-test: powershell/numeric_endianness_conversions/ipaddress_address_property_deprecated_endianness_check
$ip = [System.Net.IPAddress]::Parse("127.0.0.1")
$bytes = $ip.GetAddressBytes()
if ($bytes[0] -ne 127 -or $bytes[3] -ne 1) {
    Write-Host "FAIL: IPv4 byte ordering failed"
    exit 1
}
Write-Host "PASS"
exit 0
