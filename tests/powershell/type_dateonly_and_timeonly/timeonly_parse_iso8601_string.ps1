# vybe-test: powershell/type_dateonly_and_timeonly/timeonly_parse_iso8601_string
$t = [System.TimeOnly]::Parse("09:15:30", [System.Globalization.CultureInfo]::InvariantCulture)
if ($t.Hour -ne 9 -or $t.Minute -ne 15 -or $t.Second -ne 30) { Write-Host "FAIL: TimeOnly Parse failed"; exit 1 }
Write-Host "PASS"; exit 0
