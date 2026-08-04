# vybe-test: powershell/command_invocation/command_with_subexpression
if ((Write-Output $(1 + 1)) -eq 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
