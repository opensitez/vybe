# vybe-test: powershell/script_parameters/expression_parameter
param($x)
if ($x -eq 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
