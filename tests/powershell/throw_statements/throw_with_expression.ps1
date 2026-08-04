# vybe-test: powershell/throw_statements/throw_with_expression
$thrown = $false
try { throw (1 + 1) } catch { if ($_.Exception.Message -eq '2') { $thrown = $true } }
if ($thrown) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
