# vybe-test: powershell/boolean_expressions/compare_boolean
if (($true -eq $true) -and ($true -ne $false)) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
