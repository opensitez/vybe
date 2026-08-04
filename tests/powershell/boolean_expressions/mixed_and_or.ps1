# vybe-test: powershell/boolean_expressions/mixed_and_or
if ($false -or ($true -and $true)) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
