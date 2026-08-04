# vybe-test: powershell/function_scope/local_variable
function Test-Func { $x = 1; return $x }
if ((Test-Func) -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
