# vybe-test: powershell/type_ipaddress_parsing_and_masks/parse_ipv4_standard_dotted
$ip = [System.Net.IPAddress]::Parse("192.168.1.1")
if ($ip.ToString() -ne "192.168.1.1") {
    Write-Host "FAIL: IPv4 parsing failed"
    exit 1
}
Write-Host "PASS"
exit 0
