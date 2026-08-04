# vybe-test: powershell/function_metadata/parameter_validation_help
function Test-Func { [CmdletBinding()] param([ValidateNotNull()]$x) Write-Output $x }
if ((Test-Func -x 'PASS') -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
