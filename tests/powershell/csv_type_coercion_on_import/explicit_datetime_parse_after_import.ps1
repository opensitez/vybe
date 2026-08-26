# vybe-test: powershell/csv_type_coercion_on_import/explicit_datetime_parse_after_import
$csv = @"
Event,Timestamp
Login,2026-08-26T12:00:00Z
"@
$row = $csv | ConvertFrom-Csv
$dt = [datetime]::Parse($row.Timestamp)
if ($dt.Year -ne 2026 -or $dt.Month -ne 8 -or $dt.Day -ne 26) {
    Write-Host "FAIL: Explicit DateTime parse after import failed"
    exit 1
}
Write-Host "PASS"
exit 0
