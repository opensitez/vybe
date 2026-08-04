# vybe-test: powershell/command_quoting/argument_with_spaces
if ((Write-Output "Hello World") -eq 'Hello World') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
