# vybe-test: powershell/null_handling/compare_null
if ($null -eq $null) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
