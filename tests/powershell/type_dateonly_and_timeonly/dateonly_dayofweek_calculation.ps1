# vybe-test: powershell/type_dateonly_and_timeonly/dateonly_dayofweek_calculation
$d = [System.DateOnly]::new(2026, 8, 26) # Wednesday
if ($d.DayOfWeek -ne [System.DayOfWeek]::Wednesday) { Write-Host "FAIL: DateOnly DayOfWeek failed"; exit 1 }
Write-Host "PASS"; exit 0
