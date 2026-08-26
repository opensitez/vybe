# vybe-test: powershell/json_nested_payload_depth/depth_with_empty_hashtable_at_deep_level
$tree = @{ L1 = @{ Map = @{ Key = "Val" } } }
$json = $tree | ConvertTo-Json -Depth 3
$recovered = $json | ConvertFrom-Json
if ($recovered.L1.Map.Key -ne "Val") {
    Write-Host "FAIL: Deep hashtable serialization failed"
    exit 1
}
Write-Host "PASS"
exit 0
