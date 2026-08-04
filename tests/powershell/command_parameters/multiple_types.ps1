# vybe-test: powershell/command_parameters/multiple_types
function Test-Func { param($x,[int]$y); return $x + $y }
if ((Test-Func '1' 2) -eq '12') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
