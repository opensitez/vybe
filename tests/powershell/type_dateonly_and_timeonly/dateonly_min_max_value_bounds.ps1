# vybe-test: powershell/type_dateonly_and_timeonly/dateonly_min_max_value_bounds
$min = [System.DateOnly]::MinValue
$max = [System.DateOnly]::MaxValue
if ($min.Year -ne 1 -or $max.Year -ne 9999) { Write-Host "FAIL: DateOnly Min/Max bounds failed"; exit 1 }
Write-Host "PASS"; exit 0
