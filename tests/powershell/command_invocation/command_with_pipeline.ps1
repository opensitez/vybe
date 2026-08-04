# vybe-test: powershell/command_invocation/command_with_pipeline
if ((Write-Output 'PASS' | ForEach-Object { $_ }) -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
