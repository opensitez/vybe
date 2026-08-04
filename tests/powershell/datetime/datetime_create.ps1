# vybe-test: powershell/datetime/datetime_create
$d = [DateTime]::new(2024, 3, 15)
if ($d.Year  -ne 2024) { Write-Host "FAIL: year";  exit 1 }
if ($d.Month -ne 3)    { Write-Host "FAIL: month"; exit 1 }
if ($d.Day   -ne 15)   { Write-Host "FAIL: day";   exit 1 }
Write-Host "PASS"
exit 0
