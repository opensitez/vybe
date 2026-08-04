# vybe-test: powershell/property_access/nested_property
$obj = [pscustomobject]@{ Inner = [pscustomobject]@{ Value = 1 } }
if ($obj.Inner.Value -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
