# vybe-test: powershell/boolean_expressions/boolean_in_if
if (1 -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
