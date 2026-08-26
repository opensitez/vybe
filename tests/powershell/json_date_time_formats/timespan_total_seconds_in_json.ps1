# vybe-test: powershell/json_date_time_formats/timespan_total_seconds_in_json
$ts = [timespan]::FromMinutes(5)
$json = @{ Duration = $ts.TotalSeconds } | ConvertTo-Json
$obj = $json | ConvertFrom-Json
if ($obj.Duration -ne 300) {
    Write-Host "FAIL: TimeSpan total seconds in JSON failed, got $($obj.Duration)"
    exit 1
}
Write-Host "PASS"
exit 0
