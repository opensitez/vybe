# vybe-test: powershell/type_dateonly_and_timeonly/dateonly_addyears_calculation
$d = [System.DateOnly]::new(2026, 8, 26)
$next = $d.AddYears(4)
if ($next.Year -ne 2030) { Write-Host "FAIL: DateOnly AddYears failed"; exit 1 }
Write-Host "PASS"; exit 0
