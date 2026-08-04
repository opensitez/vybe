# vybe-test: powershell/function_metadata/parameter_attribute
function Test-Func { [CmdletBinding()] param([Parameter(Mandatory=$true)]$x) Write-Output $x }
if ((Test-Func -x 'PASS') -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
