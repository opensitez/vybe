# vybe-test: powershell/function_metadata/validate_not_null_or_empty
function Test-Func { [CmdletBinding()] param([ValidateNotNullOrEmpty()]$x) Write-Output $x }
if ((Test-Func -x 'PASS') -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
