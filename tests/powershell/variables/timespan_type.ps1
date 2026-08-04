# vybe-test: powershell/variables/timespan_type
$span = New-TimeSpan -Hours 2 -Minutes 30
$minutes = $span.TotalMinutes
if ($minutes -ne 150) {
    Write-Host "FAIL: expected 150 minutes, got $minutes"
    exit 1
}
Write-Host "PASS"
exit 0
