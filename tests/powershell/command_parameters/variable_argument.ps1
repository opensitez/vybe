# vybe-test: powershell/command_parameters/variable_argument
function Test-Func { param($x); return $x }
$val = 5
if ((Test-Func $val) -eq 5) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
