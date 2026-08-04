# vybe-test: powershell/script_parameters/type_parameter
param([int]$x)
if ($x -eq 5) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
