# vybe-test: powershell/throw_statements/throw_after_return
$thrown = $false
try { return; throw 'ERROR' } catch { $thrown = $true }
if (-not $thrown) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
