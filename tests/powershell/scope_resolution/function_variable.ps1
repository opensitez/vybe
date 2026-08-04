# vybe-test: powershell/scope_resolution/function_variable
function Test-Func { $y = 3; return $y }
if ((Test-Func) -eq 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
