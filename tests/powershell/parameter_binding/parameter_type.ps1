# vybe-test: powershell/parameter_binding/parameter_type
function Test-Func { param([int]$x); return $x }
if ((Test-Func -x 5) -eq 5) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
