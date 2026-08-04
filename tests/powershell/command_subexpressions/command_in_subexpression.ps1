# vybe-test: powershell/command_subexpressions/command_in_subexpression
if ("$(Get-Command Write-Output).Name" -eq 'Write-Output') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
