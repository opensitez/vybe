# vybe-test: powershell/numeric_endianness_conversions/network_to_host_order_int32
$hostOrder = 0x12345678
$netOrder = [System.Net.IPAddress]::HostToNetworkOrder($hostOrder)
$backToHost = [System.Net.IPAddress]::NetworkToHostOrder($netOrder)
if ($hostOrder -ne $backToHost) {
    Write-Host "FAIL: HostToNetworkOrder / NetworkToHostOrder roundtrip failed"
    exit 1
}
Write-Host "PASS"
exit 0
