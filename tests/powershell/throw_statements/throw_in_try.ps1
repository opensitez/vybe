# vybe-test: powershell/throw_statements/throw_in_try
$thrown = $false
try { throw 'ERROR' } catch { $thrown = $true }
if ($thrown) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
