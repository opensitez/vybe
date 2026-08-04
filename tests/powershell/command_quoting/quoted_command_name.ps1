# vybe-test: powershell/command_quoting/quoted_command_name
& 'Write-Output' 'PASS' | ForEach-Object { if ($_ -eq 'PASS') { Write-Host 'PASS'; exit 0 } }
Write-Host 'FAIL'
exit 1
