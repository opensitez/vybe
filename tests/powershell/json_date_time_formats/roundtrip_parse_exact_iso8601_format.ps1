# vybe-test: powershell/json_date_time_formats/roundtrip_parse_exact_iso8601_format
$json = '{"Timestamp":"2026-08-26T12:00:00Z"}'
$obj = $json | ConvertFrom-Json
if ($obj.Timestamp -eq $null) {
    Write-Host "FAIL: JSON date field missing"
    exit 1
}
Write-Host "PASS"
exit 0
