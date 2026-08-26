# vybe-test: powershell/type_ipaddress_parsing_and_masks/parse_ipv6_standard_hex
$ip = [System.Net.IPAddress]::Parse("2001:0db8:85a3:0000:0000:8a2e:0370:7334")
if ($ip.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetworkV6) {
    Write-Host "FAIL: IPv6 parsing failed"
    exit 1
}
Write-Host "PASS"
exit 0
