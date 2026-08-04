# vybe-test: powershell/command_parameters/default_argument
function Test-Func { param($x = 1,$y = 2); return $x + $y }
if ((Test-Func) -eq 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
