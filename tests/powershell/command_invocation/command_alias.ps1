# vybe-test: powershell/command_invocation/command_alias
Set-Alias out Write-Output
if ((out 'PASS') -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
