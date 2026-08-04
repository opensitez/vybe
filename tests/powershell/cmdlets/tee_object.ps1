# vybe-test: powershell/cmdlets/tee_object
$log = @()
$result = 1..5 | Tee-Object -Variable log | Measure-Object -Sum
if ($result.Sum -ne 15) { Write-Host "FAIL: sum $($result.Sum)"; exit 1 }
if ($log.Count -ne 5)   { Write-Host "FAIL: log count $($log.Count)"; exit 1 }
Write-Host "PASS"
exit 0
