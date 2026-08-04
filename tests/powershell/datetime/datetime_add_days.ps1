# vybe-test: powershell/datetime/datetime_add_days
$d = [DateTime]::new(2024, 1, 28)
$d2 = $d.AddDays(5)
if ($d2.Month -ne 2) { Write-Host "FAIL: month should be 2, got $($d2.Month)"; exit 1 }
if ($d2.Day   -ne 2) { Write-Host "FAIL: day should be 2, got $($d2.Day)";   exit 1 }
Write-Host "PASS"
exit 0
