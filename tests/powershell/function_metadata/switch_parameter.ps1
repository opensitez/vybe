# vybe-test: powershell/function_metadata/switch_parameter
function Test-Func { [CmdletBinding()] param([switch]$Flag) if ($Flag) { Write-Output 'PASS' } }
if ((Test-Func -Flag) -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
