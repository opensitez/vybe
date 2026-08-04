# vybe-test: powershell/scope_resolution/nested_function_scope
$y = 1
function Test-Func { $y = 2; return $y }
if ((Test-Func) -eq 2 -and $y -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
