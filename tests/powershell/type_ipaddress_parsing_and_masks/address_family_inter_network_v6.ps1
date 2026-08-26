# vybe-test: powershell/type_ipaddress_parsing_and_masks/address_family_inter_network_v6
$ip = [System.Net.IPAddress]::Parse("fe80::1")
if ($ip.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetworkV6) {
    Write-Host "FAIL: AddressFamily InterNetworkV6 expected"
    exit 1
}
Write-Host "PASS"
exit 0
