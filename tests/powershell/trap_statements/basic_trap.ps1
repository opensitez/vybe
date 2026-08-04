# vybe-test: powershell/trap_statements/basic_trap
$caught = $false
trap { $caught = $true; continue }
throw 'ERR'
if ($caught) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
