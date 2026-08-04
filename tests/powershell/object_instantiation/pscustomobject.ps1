# vybe-test: powershell/object_instantiation/pscustomobject
$obj = [pscustomobject]@{ A = 1 }
if ($obj.A -ne 1) { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
