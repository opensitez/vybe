# vybe-test: powershell/type_dateonly_and_timeonly/dateonly_parse_iso8601_string
$d = [System.DateOnly]::Parse("2026-12-31", [System.Globalization.CultureInfo]::InvariantCulture)
if ($d.Year -ne 2026 -or $d.Month -ne 12 -or $d.Day -ne 31) { Write-Host "FAIL: DateOnly Parse failed"; exit 1 }
Write-Host "PASS"; exit 0
