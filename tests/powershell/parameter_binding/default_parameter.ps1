# vybe-test: powershell/parameter_binding/default_parameter
function Test-Func { param($x = 1); return $x }
if ((Test-Func) -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
