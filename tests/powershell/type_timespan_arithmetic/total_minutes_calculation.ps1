# vybe-test: powershell/type_timespan_arithmetic/total_minutes_calculation
$ts = [timespan]::new(0, 2, 30, 0) # 2.5 hours = 150 minutes
if ($ts.TotalMinutes -ne 150.0) {
    Write-Host "FAIL: expected 150.0 total minutes, got $($ts.TotalMinutes)"
    exit 1
}
Write-Host "PASS"
exit 0
