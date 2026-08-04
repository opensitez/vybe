# vybe-test: powershell/return_values/command_return
$value = Write-Output 'PASS'
if ($value -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
