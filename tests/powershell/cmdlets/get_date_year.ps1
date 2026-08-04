# vybe-test: powershell/cmdlets/get_date_year
$date = Get-Date -Year 2020 -Month 5 -Day 15
$year = $date.Year
if ($year -ne 2020) {
    Write-Host "FAIL: expected 2020, got $year"
    exit 1
}
Write-Host "PASS"
exit 0
