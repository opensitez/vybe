# vybe-test: powershell/try_finally/finally_runs
$ran = $false
try { $ran = $true } finally { if ($ran) { Write-Host 'PASS'; exit 0 } }
Write-Host 'FAIL'
exit 1
