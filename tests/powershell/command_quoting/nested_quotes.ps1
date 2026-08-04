# vybe-test: powershell/command_quoting/nested_quotes
if ((Write-Output "She said 'PASS'") -eq "She said 'PASS'") { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
