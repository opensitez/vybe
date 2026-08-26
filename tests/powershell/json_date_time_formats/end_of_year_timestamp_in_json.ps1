# vybe-test: powershell/json_date_time_formats/end_of_year_timestamp_in_json
$json = '{"YearEnd":"2026-12-31T23:59:59Z"}'
$obj = $json | ConvertFrom-Json
$dt = [datetime]::Parse($obj.YearEnd)
if ($dt.Month -ne 12 -or $dt.Day -ne 31 -or $dt.Hour -ne 23 -or $dt.Minute -ne 59) {
    Write-Host "FAIL: Year end timestamp in JSON failed"
    exit 1
}
Write-Host "PASS"
exit 0
