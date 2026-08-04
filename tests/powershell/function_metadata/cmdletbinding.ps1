# vybe-test: powershell/function_metadata/cmdletbinding
function Test-Func { [CmdletBinding()] param() Write-Output 'PASS' }
if ((Test-Func) -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
