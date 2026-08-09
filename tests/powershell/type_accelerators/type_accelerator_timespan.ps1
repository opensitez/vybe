# vybe-test: powershell/type_accelerators/type_accelerator_timespan
$ts = [timespan]"02:30:00"
if ($ts.TotalMinutes -ne 150) {
    Write-Host "FAIL: timespan TotalMinutes expected 150, got $($ts.TotalMinutes)"
    exit 1
}
Write-Host "PASS"
exit 0
