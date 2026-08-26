# vybe-test: powershell/type_dateonly_and_timeonly/dateonly_creation_from_ymd
$d = [System.DateOnly]::new(2026, 8, 26)
if ($d.Year -ne 2026 -or $d.Month -ne 8 -or $d.Day -ne 26) { Write-Host "FAIL: DateOnly constructor failed"; exit 1 }
Write-Host "PASS"; exit 0
