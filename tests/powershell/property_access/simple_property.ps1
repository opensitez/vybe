# vybe-test: powershell/property_access/simple_property
$obj = [pscustomobject]@{ Name = 'PASS' }
if ($obj.Name -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
