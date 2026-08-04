# vybe-test: powershell/command_parameters/splat_argument
function Test-Func { param($x,$y); return $x + $y }
$args = @{ x = 1; y = 2 }
if ((Test-Func @args) -eq 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
