# vybe-test: powershell/return_values/function_return
function Test-Func { return 'PASS' }
if ((Test-Func) -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
