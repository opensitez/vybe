# vybe-test: powershell/type_dateonly_and_timeonly/timeonly_addminutes_calculation
$t = [System.TimeOnly]::new(10, 45, 0)
$next = $t.AddMinutes(30)
if ($next.Hour -ne 11 -or $next.Minute -ne 15) { Write-Host "FAIL: TimeOnly AddMinutes failed"; exit 1 }
Write-Host "PASS"; exit 0
