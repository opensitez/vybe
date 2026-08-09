# vybe-test: powershell/pscustomobject_literals/pscustomobject_property_mutation
$obj = [pscustomobject]@{ Count = 1 }
$obj.Count = 100
if ($obj.Count -ne 100) {
    Write-Host "FAIL: mutated property Count expected 100, got $($obj.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
