# vybe-test: powershell/json_nested_payload_depth/depth_with_null_values_at_deep_level
$tree = @{ L1 = @{ L2 = @{ L3 = $null } } }
$json = $tree | ConvertTo-Json -Depth 5
$recovered = $json | ConvertFrom-Json
if ($recovered.L1.L2.L3 -ne $null) {
    Write-Host "FAIL: Null values at deep level failed"
    exit 1
}
Write-Host "PASS"
exit 0
