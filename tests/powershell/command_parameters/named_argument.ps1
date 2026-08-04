# vybe-test: powershell/command_parameters/named_argument
function Test-Func { param($x,$y); return $x + $y }
if ((Test-Func -y 2 -x 1) -eq 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
