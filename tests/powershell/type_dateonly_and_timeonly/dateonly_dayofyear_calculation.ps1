# vybe-test: powershell/type_dateonly_and_timeonly/dateonly_dayofyear_calculation
$d = [System.DateOnly]::new(2026, 2, 1)
if ($d.DayOfYear -ne 32) { Write-Host "FAIL: DateOnly DayOfYear failed"; exit 1 }
Write-Host "PASS"; exit 0
