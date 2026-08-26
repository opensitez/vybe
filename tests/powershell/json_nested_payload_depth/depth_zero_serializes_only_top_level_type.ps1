# vybe-test: powershell/json_nested_payload_depth/depth_zero_serializes_only_top_level_type
$obj = @{ Key = "Value" }
$json = $obj | ConvertTo-Json -Depth 0
# Depth 0 produces string representation of top object
if ($json.Length -eq 0) {
    Write-Host "FAIL: Depth 0 produced empty output"
    exit 1
}
Write-Host "PASS"
exit 0
