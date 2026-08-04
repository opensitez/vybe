# vybe-test: powershell/command_parameters/switch_argument
function Test-Func { param([switch]$Flag); if ($Flag) { return 'PASS' } }
if ((Test-Func -Flag) -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
