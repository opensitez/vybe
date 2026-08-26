# vybe-test: powershell/json_nested_payload_depth/depth_with_compress_switch
$tree = @{ A = @{ B = 1 } }
$json = $tree | ConvertTo-Json -Depth 3 -Compress
if ($json.Contains("`n") -or -not $json.Contains('"B":1')) {
    Write-Host "FAIL: Depth with -Compress switch failed, got '$json'"
    exit 1
}
Write-Host "PASS"
exit 0
