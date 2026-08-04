# vybe-test: powershell/datetime/timespan_arithmetic
$ts1 = [TimeSpan]::new(1, 30, 0)  # 1h 30m
$ts2 = [TimeSpan]::new(0, 45, 0)  # 45m
$sum = $ts1 + $ts2
if ($sum.Hours   -ne 2)  { Write-Host "FAIL: hours"; exit 1 }
if ($sum.Minutes -ne 15) { Write-Host "FAIL: minutes"; exit 1 }
$diff = $ts1 - $ts2
if ($diff.Minutes -ne 45) { Write-Host "FAIL: diff minutes"; exit 1 }
Write-Host "PASS"
exit 0
