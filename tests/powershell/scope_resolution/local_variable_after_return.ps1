# vybe-test: powershell/scope_resolution/local_variable_after_return
function Test-Func { $x = 1; return $x; $x = 2 }
if ((Test-Func) -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
