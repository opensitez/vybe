# vybe-test: powershell/try_finally/no_error
$ran = $false
try { $ran = $true } finally { if ($ran) { Write-Host 'PASS'; exit 0 } }
Write-Host 'FAIL'
exit 1
