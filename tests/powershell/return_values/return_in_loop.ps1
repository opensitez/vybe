# vybe-test: powershell/return_values/return_in_loop
function Test-Func { for ($i=0; $i -lt 1; $i++) { return 'PASS' } }
if ((Test-Func) -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
