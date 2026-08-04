# vybe-test: powershell/type_checking/isnot_null
if ($null -isnot [int]) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
