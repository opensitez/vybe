# vybe-test: powershell/function_metadata/parameter_alias
function Test-Func { [CmdletBinding()] param([Alias('X')]$x) Write-Output $x }
if ((Test-Func -X 'PASS') -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
