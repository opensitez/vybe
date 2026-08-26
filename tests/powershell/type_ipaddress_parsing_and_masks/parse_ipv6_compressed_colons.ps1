# vybe-test: powershell/type_ipaddress_parsing_and_masks/parse_ipv6_compressed_colons
$ip = [System.Net.IPAddress]::Parse("::1")
if ($ip.ToString() -ne "::1") {
    Write-Host "FAIL: IPv6 compressed loopback parse failed"
    exit 1
}
Write-Host "PASS"
exit 0
