# vybe-test: powershell/null_handling/null_property
$obj = $null
if ($obj -eq $null) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
