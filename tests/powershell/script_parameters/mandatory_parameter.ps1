# vybe-test: powershell/script_parameters/mandatory_parameter
param([Parameter(Mandatory=$true)]$x)
if ($x -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
