# vybe-test: powershell/numeric_endianness_conversions/network_to_host_order_int64
[int64]$hostOrder = 0x0102030405060708
$netOrder = [System.Net.IPAddress]::HostToNetworkOrder($hostOrder)
$backToHost = [System.Net.IPAddress]::NetworkToHostOrder($netOrder)
if ($hostOrder -ne $backToHost) {
    Write-Host "FAIL: Int64 HostToNetworkOrder failed"
    exit 1
}
Write-Host "PASS"
exit 0
