# vybe-test: powershell/property_access/property_chaining
$obj = [pscustomobject]@{ A = [pscustomobject]@{ B = 'PASS' } }
if ($obj.A.B -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
