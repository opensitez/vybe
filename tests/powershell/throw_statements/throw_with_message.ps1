# vybe-test: powershell/throw_statements/throw_with_message
$thrown = $false
try { throw 'FAIL' } catch { if ($_.Exception.Message -eq 'FAIL') { $thrown = $true } }
if ($thrown) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
