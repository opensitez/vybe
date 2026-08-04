# vybe-test: powershell/parameter_binding/positional_parameter
function Test-Func { param($x); return $x }
if ((Test-Func 5) -eq 5) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
