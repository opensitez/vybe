# vybe-test: powershell/boolean_expressions/boolean_grouping
if (($true -or $false) -and $true) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
