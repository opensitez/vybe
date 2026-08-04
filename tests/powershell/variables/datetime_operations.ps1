# vybe-test: powershell/variables/datetime_operations
$date = Get-Date "2023-01-15"
$date = $date.AddDays(10)
$month = $date.Month
if ($month -ne 1) {
    Write-Host "FAIL: expected month 1, got $month"
    exit 1
}
$day = $date.Day
if ($day -ne 25) {
    Write-Host "FAIL: expected day 25, got $day"
    exit 1
}
Write-Host "PASS"
exit 0
