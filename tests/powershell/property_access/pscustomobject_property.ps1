# vybe-test: powershell/property_access/pscustomobject_property
$obj = [pscustomobject]@{ Value = 10 }
if ($obj.Value -eq 10) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
