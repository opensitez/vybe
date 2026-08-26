# vybe-test: powershell/type_timespan_arithmetic/from_days_hours_minutes
$ts = [timespan]::new(2, 3, 4, 5) # 2 days, 3 hours, 4 mins, 5 secs
if ($ts.Days -ne 2 -or $ts.Hours -ne 3 -or $ts.Minutes -ne 4 -or $ts.Seconds -ne 5) {
    Write-Host "FAIL: TimeSpan components mismatched"
    exit 1
}
Write-Host "PASS"
exit 0
