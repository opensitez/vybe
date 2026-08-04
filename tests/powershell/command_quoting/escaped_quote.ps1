# vybe-test: powershell/command_quoting/escaped_quote
if ((Write-Output "He said \"PASS\"") -eq 'He said "PASS"') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
