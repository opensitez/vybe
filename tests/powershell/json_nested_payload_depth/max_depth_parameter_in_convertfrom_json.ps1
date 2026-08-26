# vybe-test: powershell/json_nested_payload_depth/max_depth_parameter_in_convertfrom_json
$json = '{"a":{"b":{"c":{"d":123}}}}'
$obj = ConvertFrom-Json -InputObject $json -Depth 10
if ($obj.a.b.c.d -ne 123) {
    Write-Host "FAIL: ConvertFrom-Json with -Depth 10 failed"
    exit 1
}
Write-Host "PASS"
exit 0
