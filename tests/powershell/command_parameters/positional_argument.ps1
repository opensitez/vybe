# vybe-test: powershell/command_parameters/positional_argument
function Test-Func { param($x,$y); return $x + $y }
if ((Test-Func 1 2) -eq 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
