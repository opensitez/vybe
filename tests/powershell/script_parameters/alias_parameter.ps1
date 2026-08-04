# vybe-test: powershell/script_parameters/alias_parameter
param([Alias('X')]$x)
if ($x -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
