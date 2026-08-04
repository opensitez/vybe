# vybe-test: powershell/pipeline/measure_object_stats
$nums = 1..10
$stats = $nums | Measure-Object -Sum -Average -Minimum -Maximum
if ($stats.Sum     -ne 55)  { Write-Host "FAIL: Sum";     exit 1 }
if ($stats.Average -ne 5.5) { Write-Host "FAIL: Average"; exit 1 }
if ($stats.Minimum -ne 1)   { Write-Host "FAIL: Min";     exit 1 }
if ($stats.Maximum -ne 10)  { Write-Host "FAIL: Max";     exit 1 }
Write-Host "PASS"
exit 0
