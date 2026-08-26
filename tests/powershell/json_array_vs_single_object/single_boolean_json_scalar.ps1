# vybe-test: powershell/json_array_vs_single_object/single_boolean_json_scalar
$json = 'true'
$val = ConvertFrom-Json -InputObject $json
if ($val -ne $true) {
    Write-Host "FAIL: Single boolean scalar JSON failed"
    exit 1
}
Write-Host "PASS"
exit 0
