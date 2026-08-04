# vybe-test: powershell/parameter_binding/named_parameter
function Test-Func { param($x); return $x }
if ((Test-Func -x 7) -eq 7) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
