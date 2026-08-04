# vybe-test: powershell/command_parameters/optional_parameter
function Test-Func { param($x,$y = 2); return $x + $y }
if ((Test-Func 1) -eq 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
