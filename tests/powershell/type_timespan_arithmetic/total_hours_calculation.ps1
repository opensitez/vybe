# vybe-test: powershell/type_timespan_arithmetic/total_hours_calculation
$ts = [timespan]::new(1, 12, 0, 0) # 1.5 days = 36 hours
if ($ts.TotalHours -ne 36.0) {
    Write-Host "FAIL: expected 36.0 total hours, got $($ts.TotalHours)"
    exit 1
}
Write-Host "PASS"
exit 0
