# vybe-test: powershell/json_custom_object_roundtrip/custom_object_with_ipaddress_property
$ip = [System.Net.IPAddress]::Parse("192.168.1.50")
$orig = [pscustomobject]@{ HostIp = $ip.ToString() }
$json = $orig | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
if ($recovered.HostIp -ne "192.168.1.50") {
    Write-Host "FAIL: IPAddress property roundtrip failed"
    exit 1
}
Write-Host "PASS"
exit 0
