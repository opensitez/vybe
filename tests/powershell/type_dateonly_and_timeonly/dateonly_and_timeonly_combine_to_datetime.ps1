# vybe-test: powershell/type_dateonly_and_timeonly/dateonly_and_timeonly_combine_to_datetime
$d = [System.DateOnly]::new(2026, 8, 26)
$t = [System.TimeOnly]::new(10, 0, 0)
$dt = $d.ToDateTime($t)
if ($dt.Year -ne 2026 -or $dt.Hour -ne 10) { Write-Host "FAIL: DateOnly ToDateTime failed"; exit 1 }
Write-Host "PASS"; exit 0
