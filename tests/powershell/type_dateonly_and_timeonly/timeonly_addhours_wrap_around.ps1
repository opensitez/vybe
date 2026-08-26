# vybe-test: powershell/type_dateonly_and_timeonly/timeonly_addhours_wrap_around
$t = [System.TimeOnly]::new(22, 0, 0)
$next = $t.AddHours(4)
if ($next.Hour -ne 2) { Write-Host "FAIL: TimeOnly AddHours wrap around failed"; exit 1 }
Write-Host "PASS"; exit 0
