# vybe-test: powershell/trap_statements/try_trap
$caught = $false
trap { $caught = $true; continue }
try { throw 'ERR' } catch { }
if ($caught) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
