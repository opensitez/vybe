# vybe-test: powershell/try_finally/with_error
$ran = $false
try { throw 'ERR' } finally { $ran = $true }
if ($ran) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
