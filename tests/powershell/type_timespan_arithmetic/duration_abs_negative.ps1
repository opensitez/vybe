# vybe-test: powershell/type_timespan_arithmetic/duration_abs_negative
$ts = [timespan]::FromMinutes(-45)
$abs = $ts.Duration()
if ($abs.TotalMinutes -ne 45.0) {
    Write-Host "FAIL: Duration() expected 45, got $($abs.TotalMinutes)"
    exit 1
}
Write-Host "PASS"
exit 0
