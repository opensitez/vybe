# vybe-test: powershell/type_dateonly_and_timeonly/dateonly_from_datetime
$dt = [datetime]::ParseExact("2026-08-26 14:30:00", "yyyy-MM-dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture)
$d = [System.DateOnly]::FromDateTime($dt)
if ($d.Year -ne 2026 -or $d.Month -ne 8 -or $d.Day -ne 26) { Write-Host "FAIL: DateOnly FromDateTime failed"; exit 1 }
Write-Host "PASS"; exit 0
