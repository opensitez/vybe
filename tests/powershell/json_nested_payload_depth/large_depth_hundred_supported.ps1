# vybe-test: powershell/json_nested_payload_depth/large_depth_hundred_supported
$obj = @{ Data = 100 }
$json = $obj | ConvertTo-Json -Depth 100
$recovered = $json | ConvertFrom-Json
if ($recovered.Data -ne 100) {
    Write-Host "FAIL: ConvertTo-Json with Depth 100 failed"
    exit 1
}
Write-Host "PASS"
exit 0
