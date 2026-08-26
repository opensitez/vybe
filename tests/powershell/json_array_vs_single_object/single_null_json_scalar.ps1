# vybe-test: powershell/json_array_vs_single_object/single_null_json_scalar
$json = 'null'
$val = ConvertFrom-Json -InputObject $json
if ($val -ne $null) {
    Write-Host "FAIL: Single null scalar JSON failed"
    exit 1
}
Write-Host "PASS"
exit 0
