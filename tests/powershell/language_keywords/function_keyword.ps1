# vybe-test: powershell/language_keywords/function_keyword
function Test-Func { return 'PASS' }
if ((Test-Func) -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
