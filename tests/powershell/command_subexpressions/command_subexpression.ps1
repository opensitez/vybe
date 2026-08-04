# vybe-test: powershell/command_subexpressions/command_subexpression
if ("$(Write-Output 'PASS')" -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
