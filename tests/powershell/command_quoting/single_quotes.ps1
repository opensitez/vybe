# vybe-test: powershell/command_quoting/single_quotes
if ((Write-Output 'PASS') -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
