# vybe-test: powershell/function_metadata/parameter_help
function Test-Func { [CmdletBinding()] param([Parameter(HelpMessage='help')]$x) Write-Output 'PASS' }
if ((Test-Func -x 1) -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
