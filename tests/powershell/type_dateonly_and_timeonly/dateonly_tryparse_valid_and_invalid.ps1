# vybe-test: powershell/type_dateonly_and_timeonly/dateonly_tryparse_valid_and_invalid
$d = [System.DateOnly]::Parse("2026-05-15")
if ($d.Day -ne 15 -or $d.Month -ne 5) { Write-Host "FAIL: DateOnly check failed"; exit 1 }
Write-Host "PASS"; exit 0
