# vybe-test: powershell/throw_statements/throw_in_scriptblock
$thrown = $false
$sb = { throw 'ERROR' }
try { & $sb } catch { $thrown = $true }
if ($thrown) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
