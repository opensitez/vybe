# vybe-test: powershell/type_accelerators/type_accelerator_ipaddress
$ip = [ipaddress]"192.168.1.100"
if ($ip.IPAddressToString -ne "192.168.1.100") {
    Write-Host "FAIL: IP string expected 192.168.1.100, got $($ip.IPAddressToString)"
    exit 1
}
if ($ip.AddressFamily -ne "InterNetwork") {
    Write-Host "FAIL: AddressFamily expected InterNetwork, got $($ip.AddressFamily)"
    exit 1
}
Write-Host "PASS"
exit 0
