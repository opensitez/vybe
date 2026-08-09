# vybe-test: powershell/pscustomobject_literals/pscustomobject_property_access
$obj = [pscustomobject]@{ Item = "Value" }
$propName = "Item"
if ($obj.$propName -ne "Value") {
    Write-Host "FAIL: dynamic property access expected Value, got $($obj.$propName)"
    exit 1
}
Write-Host "PASS"
exit 0
