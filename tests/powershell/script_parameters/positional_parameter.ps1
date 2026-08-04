# vybe-test: powershell/script_parameters/positional_parameter
param($x,$y)
if ($x -eq 1 -and $y -eq 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
