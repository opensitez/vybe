# vybe-test: powershell/property_access/parenthesized_property
$obj = [pscustomobject]@{ Value = 'PASS' }
if ( ($obj).Value -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
