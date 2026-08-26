# vybe-test: powershell/json_array_vs_single_object/array_json_returns_object_array
$json = '[{"id":1},{"id":2}]'
$arr = ConvertFrom-Json -InputObject $json
if ($arr.Count -ne 2 -or $arr[0].id -ne 1 -or $arr[1].id -ne 2) {
    Write-Host "FAIL: Array JSON should produce array of objects"
    exit 1
}
Write-Host "PASS"
exit 0
