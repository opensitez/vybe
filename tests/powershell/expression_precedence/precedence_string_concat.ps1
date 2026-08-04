# vybe-test: powershell/expression_precedence/precedence_string_concat
if ('A' + 'B' -eq 'AB') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
