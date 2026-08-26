# vybe-test: powershell/type_dateonly_and_timeonly/timeonly_from_timespan
$ts = [timespan]::FromHours(15.5)
$t = [System.TimeOnly]::FromTimeSpan($ts)
if ($t.Hour -ne 15 -or $t.Minute -ne 30) { Write-Host "FAIL: TimeOnly FromTimeSpan failed"; exit 1 }
Write-Host "PASS"; exit 0
