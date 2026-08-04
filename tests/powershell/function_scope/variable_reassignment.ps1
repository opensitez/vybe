# vybe-test: powershell/function_scope/variable_reassignment
$x = 1
function Test-Func { $x = 2 }
Test-Func
if ($x -eq 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
