# vybe-test: powershell/throw_statements/throw_in_if
$thrown = $false
try { if ($true) { throw 'ERROR' } } catch { $thrown = $true }
if ($thrown) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
