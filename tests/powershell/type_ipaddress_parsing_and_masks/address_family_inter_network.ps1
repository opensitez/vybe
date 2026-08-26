# vybe-test: powershell/type_ipaddress_parsing_and_masks/address_family_inter_network
$ip = [System.Net.IPAddress]::Parse("172.16.0.5")
if ($ip.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) {
    Write-Host "FAIL: AddressFamily InterNetwork expected"
    exit 1
}
Write-Host "PASS"
exit 0
