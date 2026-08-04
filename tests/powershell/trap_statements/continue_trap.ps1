# vybe-test: powershell/trap_statements/continue_trap
$caught = $false
trap { $caught = $true; continue }
1..1 | ForEach-Object { throw 'ERR' }
if ($caught) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
