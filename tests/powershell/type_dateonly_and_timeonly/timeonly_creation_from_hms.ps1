# vybe-test: powershell/type_dateonly_and_timeonly/timeonly_creation_from_hms
$t = [System.TimeOnly]::new(14, 30, 45)
if ($t.Hour -ne 14 -or $t.Minute -ne 30 -or $t.Second -ne 45) { Write-Host "FAIL: TimeOnly constructor failed"; exit 1 }
Write-Host "PASS"; exit 0
