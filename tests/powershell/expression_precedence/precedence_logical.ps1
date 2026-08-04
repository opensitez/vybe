# vybe-test: powershell/expression_precedence/precedence_logical
if ((1 -eq 1) -or $false) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
