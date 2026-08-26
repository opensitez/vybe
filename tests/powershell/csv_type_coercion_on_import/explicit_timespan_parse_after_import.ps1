# vybe-test: powershell/csv_type_coercion_on_import/explicit_timespan_parse_after_import
$csv = @"
Job,Duration
Backup,01:30:00
"@
$row = $csv | ConvertFrom-Csv
$ts = [timespan]::Parse($row.Duration)
if ($ts.TotalMinutes -ne 90.0) {
    Write-Host "FAIL: Explicit TimeSpan parse after import failed, got $($ts.TotalMinutes)"
    exit 1
}
Write-Host "PASS"
exit 0
