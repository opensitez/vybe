# vybe-test: powershell/dynamic_property_lookup_by_variable/dynamic_property_with_spaces_in_name
$obj = [pscustomobject]@{ "spaced property" = "spaced value" }
$p = "spaced property"
$res = $obj.$p
if ($res -ne "spaced value") {
    Write-Host "FAIL: Dynamic property with spaces in name failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
