# vybe-test: powershell/type_declarations/typed_parameter
function Test-Func { param([int]$x); return $x }
if ((Test-Func -x 2) -eq 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
