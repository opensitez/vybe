# vybe-test: powershell/trap_statements/trap_after_throw
$caught = $false
trap { $caught = $true; continue }
throw 'ERR'
if ($caught) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
