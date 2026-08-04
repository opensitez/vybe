# vybe-test: powershell/command_quoting/backtick_quote
if ((Write-Output "Line`"Break") -eq 'Line"Break') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
