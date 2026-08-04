# vybe-test: powershell/return_values/return_in_if
function Test-Func { if ($true) { return 'PASS' } }
if ((Test-Func) -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
