# vybe-test: powershell/json_nested_payload_depth/depth_preserves_special_characters_in_deep_keys
$tree = @{ "outer.key" = @{ "inner.key" = "target" } }
$json = $tree | ConvertTo-Json -Depth 3
$recovered = $json | ConvertFrom-Json
if ($recovered."outer.key"."inner.key" -ne "target") {
    Write-Host "FAIL: Special characters in deep keys failed"
    exit 1
}
Write-Host "PASS"
exit 0
