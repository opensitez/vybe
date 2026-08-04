# vybe-test: powershell/variable_declaration/uninitialized_variable
if ($null -eq $undefinedVar) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
