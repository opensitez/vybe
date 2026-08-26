# vybe-test: powershell/type_dateonly_and_timeonly/dateonly_adddays_calculation
$d = [System.DateOnly]::new(2026, 1, 30)
$next = $d.AddDays(5)
if ($next.Month -ne 2 -or $next.Day -ne 4) { Write-Host "FAIL: DateOnly AddDays failed"; exit 1 }
Write-Host "PASS"; exit 0
