# vybe-test: powershell/json_date_time_formats/datetime_array_in_json
$dates = @([datetime]::Parse("2026-01-01"), [datetime]::Parse("2026-12-31"))
$json = @{ Dates = $dates } | ConvertTo-Json
$obj = $json | ConvertFrom-Json
if ($obj.Dates.Count -ne 2) {
    Write-Host "FAIL: DateTime array in JSON failed"
    exit 1
}
Write-Host "PASS"
exit 0
