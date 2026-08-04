# vybe-test: powershell/command_invocation/command_with_arguments
if ((Write-Output 'PASS') -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
