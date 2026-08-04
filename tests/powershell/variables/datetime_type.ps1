# vybe-test: powershell/variables/datetime_type
[datetime]$date = "2023-01-15"
$year = $date.Year
if ($year -ne 2023) {
    Write-Host "FAIL: expected year 2023, got $year"
    exit 1
}
Write-Host "PASS"
exit 0
