# vybe-test: powershell/command_quoting/quoted_subexpression
if ((Write-Output "$(1 + 1)") -eq '2') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
