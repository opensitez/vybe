# vybe-test: powershell/command_invocation/command_with_variable_argument
$arg = 'PASS'
if ((Write-Output $arg) -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
