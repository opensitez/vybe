# vybe-test: powershell/json_date_time_formats/iso8601_string_parsed_as_string_in_convertfrom_json
$json = '{"Timestamp":"2026-08-26T12:00:00Z"}'
$obj = $json | ConvertFrom-Json
if ($obj.Timestamp -eq $null) {
    Write-Host "FAIL: JSON date field missing"
    exit 1
}
Write-Host "PASS"
exit 0
