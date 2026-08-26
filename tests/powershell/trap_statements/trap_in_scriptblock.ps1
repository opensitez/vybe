# vybe-test: powershell/trap_statements/trap_in_scriptblock
$caught = $false
trap { $script:caught = $true; continue }
& { throw 'ERR' }
if ($caught) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
