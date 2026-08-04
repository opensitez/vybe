# vybe-test: powershell/parameter_binding/positional_and_named
function Test-Func { param($x,$y); return $x + $y }
if ((Test-Func 1 -y 2) -eq 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
