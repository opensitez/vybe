# vybe-test: powershell/expression_precedence/precedence_hashtable
if ((@{a=1}.a) -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
