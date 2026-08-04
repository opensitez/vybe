# vybe-test: powershell/return_values/return_without_value
function Test-Func { return }
if ((Test-Func) -eq $null) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
