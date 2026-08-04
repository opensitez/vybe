# vybe-test: powershell/script_parameters/parameter_help
param([Parameter(HelpMessage='help')]$x)
if ($x -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
