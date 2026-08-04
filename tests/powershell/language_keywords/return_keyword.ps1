# vybe-test: powershell/language_keywords/return_keyword
function Test-Func { return 'PASS'; Write-Host 'FAIL' }
if ((Test-Func) -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
