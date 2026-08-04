# vybe-test: powershell/command_invocation/simple_command
if ((Get-Command Write-Output).Name -eq 'Write-Output') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
