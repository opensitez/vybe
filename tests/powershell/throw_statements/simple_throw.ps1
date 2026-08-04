# vybe-test: powershell/throw_statements/simple_throw
$thrown = $false
try { throw 'ERROR' } catch { $thrown = $true }
if ($thrown) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
