# vybe-test: powershell/datetime/datetime_difference_timespan
$d1 = [DateTime]::new(2024, 1, 1)
$d2 = [DateTime]::new(2024, 1, 11)
$span = $d2 - $d1
if ($span.Days -ne 10) {
    Write-Host "FAIL: expected 10 days, got $($span.Days)"
    exit 1
}
Write-Host "PASS"
exit 0
