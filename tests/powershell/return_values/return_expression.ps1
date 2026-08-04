# vybe-test: powershell/return_values/return_expression
function Test-Func { return (1 + 1) }
if ((Test-Func) -eq 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
