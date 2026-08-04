# vybe-test: powershell/command_quoting/variable_in_quotes
$name = 'PASS'
if ((Write-Output "$name") -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
