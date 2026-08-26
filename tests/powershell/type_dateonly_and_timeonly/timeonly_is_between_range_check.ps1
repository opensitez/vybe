# vybe-test: powershell/type_dateonly_and_timeonly/timeonly_is_between_range_check
$t = [System.TimeOnly]::new(12, 0, 0)
$start = [System.TimeOnly]::new(9, 0, 0)
$end = [System.TimeOnly]::new(17, 0, 0)
if (-not $t.IsBetween($start, $end)) { Write-Host "FAIL: TimeOnly IsBetween failed"; exit 1 }
Write-Host "PASS"; exit 0
