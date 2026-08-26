# vybe-test: powershell/dynamic_property_lookup_by_variable/dynamic_property_non_existent_returns_null
$obj = [pscustomobject]@{ Valid = 1 }
$prop = "MissingProp"
$res = $obj.$prop
if ($res -ne $null) {
    Write-Host "FAIL: Non-existent dynamic property should return null, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
