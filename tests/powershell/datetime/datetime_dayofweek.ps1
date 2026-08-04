# vybe-test: powershell/datetime/datetime_dayofweek
$d = [DateTime]::new(2024, 1, 1)   # Monday
if ($d.DayOfWeek -ne [DayOfWeek]::Monday) {
    Write-Host "FAIL: 2024-01-01 should be Monday, got $($d.DayOfWeek)"
    exit 1
}
$weekend = [DateTime]::new(2024, 1, 6)  # Saturday
if ($weekend.DayOfWeek -ne [DayOfWeek]::Saturday) {
    Write-Host "FAIL: expected Saturday, got $($weekend.DayOfWeek)"
    exit 1
}
Write-Host "PASS"
exit 0
