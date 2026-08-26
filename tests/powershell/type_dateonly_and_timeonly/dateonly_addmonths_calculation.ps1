# vybe-test: powershell/type_dateonly_and_timeonly/dateonly_addmonths_calculation
$d = [System.DateOnly]::new(2026, 1, 15)
$next = $d.AddMonths(3)
if ($next.Month -ne 4 -or $next.Day -ne 15) { Write-Host "FAIL: DateOnly AddMonths failed"; exit 1 }
Write-Host "PASS"; exit 0
