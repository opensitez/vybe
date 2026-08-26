# vybe-test: powershell/type_timespan_arithmetic/total_seconds_calculation
$ts = [timespan]::new(0, 0, 5, 30) # 5 mins 30 secs = 330 secs
if ($ts.TotalSeconds -ne 330.0) {
    Write-Host "FAIL: expected 330.0 total seconds, got $($ts.TotalSeconds)"
    exit 1
}
Write-Host "PASS"
exit 0
