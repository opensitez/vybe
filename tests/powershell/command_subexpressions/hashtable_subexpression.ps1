# vybe-test: powershell/command_subexpressions/hashtable_subexpression
$h = @{ a = 1 }
if ("$($h.a)" -eq '1') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
