# vybe-test: powershell/script_parameters/switch_parameter
param([switch]$Flag)
if ($Flag) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
