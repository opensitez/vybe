# vybe-test: powershell/trap_statements/trap_with_message
$caught = $false
trap { if ($_.Exception.Message -eq 'ERR') { $caught = $true }; continue }
throw 'ERR'
if ($caught) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
