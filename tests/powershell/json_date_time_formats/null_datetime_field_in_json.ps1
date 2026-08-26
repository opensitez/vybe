# vybe-test: powershell/json_date_time_formats/null_datetime_field_in_json
$json = '{"ExpireDate":null}'
$obj = $json | ConvertFrom-Json
if ($obj.ExpireDate -ne $null) {
    Write-Host "FAIL: Null date field in JSON failed"
    exit 1
}
Write-Host "PASS"
exit 0
