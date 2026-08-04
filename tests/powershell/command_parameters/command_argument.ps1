# vybe-test: powershell/command_parameters/command_argument
function Test-Func { param($x); return $x }
if ((Test-Func (Write-Output 5)) -eq 5) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
