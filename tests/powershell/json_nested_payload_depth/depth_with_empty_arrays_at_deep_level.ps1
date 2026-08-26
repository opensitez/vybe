# vybe-test: powershell/json_nested_payload_depth/depth_with_empty_arrays_at_deep_level
$tree = @{ L1 = @{ Items = @() } }
$json = $tree | ConvertTo-Json -Depth 3
$recovered = $json | ConvertFrom-Json
if ($recovered.L1.Items.Count -ne 0) {
    Write-Host "FAIL: Empty array at deep level failed"
    exit 1
}
Write-Host "PASS"
exit 0
