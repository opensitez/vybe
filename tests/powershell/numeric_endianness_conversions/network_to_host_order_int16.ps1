# vybe-test: powershell/numeric_endianness_conversions/network_to_host_order_int16
[int16]$hostOrder = 0x1234
$netOrder = [System.Net.IPAddress]::HostToNetworkOrder($hostOrder)
$backToHost = [System.Net.IPAddress]::NetworkToHostOrder($netOrder)
if ($hostOrder -ne $backToHost) {
    Write-Host "FAIL: Int16 HostToNetworkOrder failed"
    exit 1
}
Write-Host "PASS"
exit 0
