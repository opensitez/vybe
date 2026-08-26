# vybe-test: powershell/json_nested_payload_depth/depth_one_flattens_nested_properties
$obj = @{ User = @{ Name = "Alice" } }
$json = $obj | ConvertTo-Json -Depth 1
if ($json -eq $null -or $json.Length -eq 0) {
    Write-Host "FAIL: ConvertTo-Json depth 1 failed"
    exit 1
}
Write-Host "PASS"
exit 0
