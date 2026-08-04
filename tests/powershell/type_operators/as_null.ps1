# vybe-test: powershell/type_operators/as_null
$value = $null -as [string]
if ($value -eq $null) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
