# vybe-test: powershell/null_handling/null_function_return
function Test-Func { return $null }
if ((Test-Func) -eq $null) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
