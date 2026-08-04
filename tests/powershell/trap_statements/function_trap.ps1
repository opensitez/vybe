# vybe-test: powershell/trap_statements/function_trap
$caught = $false
trap { $caught = $true; continue }
function Test-Func { throw 'ERR' }
Test-Func
if ($caught) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
