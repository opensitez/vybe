# vybe-test: powershell/command_subexpressions/variable_subexpression
$x = 5
if ("$($x)" -eq '5') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
