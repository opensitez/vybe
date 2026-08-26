# vybe-test: powershell/dynamic_property_lookup_by_variable/dynamic_property_with_special_characters_in_name
$obj = [pscustomobject]@{ "my.special#prop" = "found" }
$p = "my.special#prop"
$res = $obj.$p
if ($res -ne "found") {
    Write-Host "FAIL: Dynamic property with special characters failed"
    exit 1
}
Write-Host "PASS"
exit 0
