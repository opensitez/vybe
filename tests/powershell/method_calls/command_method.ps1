# vybe-test: powershell/method_calls/command_method
if ((Get-Command Write-Output).Name -eq 'Write-Output') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
