# vybe-test: powershell/pipeline_sort_object_properties/sort_ipaddresses_ordering
$ip1 = [System.Net.IPAddress]::Parse("192.168.1.10")
$ip2 = [System.Net.IPAddress]::Parse("10.0.0.1")
$ip3 = [System.Net.IPAddress]::Parse("172.16.0.1")
$sorted = @($ip1, $ip2, $ip3 | Sort-Object { $_.ToString() })
if ($sorted[0].ToString() -ne "10.0.0.1" -or $sorted[2].ToString() -ne "192.168.1.10") {
    Write-Host "FAIL: Sort-Object IP addresses failed"
    exit 1
}
Write-Host "PASS"
exit 0
