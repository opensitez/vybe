# vybe-test: powershell/type_nullable_value_types/has_value_property_false
[System.Nullable[bool]]$b = $null
if ($b.HasValue) {
    Write-Host "FAIL: HasValue should be false"
    exit 1
}
Write-Host "PASS"
exit 0
