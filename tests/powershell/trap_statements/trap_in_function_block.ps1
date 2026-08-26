# vybe-test: powershell/trap_statements/trap_in_function_block
$caught = $false
trap { $script:caught = $true; continue }
function Test-Func { throw 'ERR' }
Test-Func
if ($caught) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
