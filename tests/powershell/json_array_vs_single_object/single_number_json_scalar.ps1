# vybe-test: powershell/json_array_vs_single_object/single_number_json_scalar
$json = '12345'
$val = ConvertFrom-Json -InputObject $json
if ($val -ne 12345) {
    Write-Host "FAIL: Single number scalar JSON failed, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
