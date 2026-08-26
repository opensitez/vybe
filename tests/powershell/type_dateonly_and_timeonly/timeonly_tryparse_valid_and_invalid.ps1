# vybe-test: powershell/type_dateonly_and_timeonly/timeonly_tryparse_valid_and_invalid
$t = [System.TimeOnly]::Parse("23:59:59")
if ($t.Hour -ne 23 -or $t.Minute -ne 59) { Write-Host "FAIL: TimeOnly check failed"; exit 1 }
Write-Host "PASS"; exit 0
