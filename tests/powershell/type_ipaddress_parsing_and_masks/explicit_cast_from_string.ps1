# vybe-test: powershell/type_ipaddress_parsing_and_masks/explicit_cast_from_string
[System.Net.IPAddress]$ip = "127.0.0.1"
if ($ip.ToString() -ne "127.0.0.1") {
    Write-Host "FAIL: cast to IPAddress failed"
    exit 1
}
Write-Host "PASS"
exit 0
