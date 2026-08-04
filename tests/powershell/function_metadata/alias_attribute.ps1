# vybe-test: powershell/function_metadata/alias_attribute
function Test-Func { [Alias('TF')] param() Write-Output 'PASS' }
if ((TF) -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
