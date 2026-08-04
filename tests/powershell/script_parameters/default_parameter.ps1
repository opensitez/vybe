# vybe-test: powershell/script_parameters/default_parameter
param($x = 1)
if ($x -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
