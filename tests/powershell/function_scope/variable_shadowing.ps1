# vybe-test: powershell/function_scope/variable_shadowing
$x = 1
function Test-Func { $x = 2; return $x }
if ((Test-Func) -eq 2 -and $x -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
