# vybe-test: powershell/return_values/return_in_scriptblock
$sb = { return 'PASS' }
if ((& $sb) -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
