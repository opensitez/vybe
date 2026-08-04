# vybe-test: powershell/function_scope/return_scope
function Test-Func { return 'PASS'; $x = 2 }
if ((Test-Func) -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
