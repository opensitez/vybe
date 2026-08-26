# vybe-test: powershell/type_nullable_value_types/has_value_property_true
$t = [type]"System.Nullable[double]"
$inst = [System.Activator]::CreateInstance($t, @(3.14))
$prop = $t.GetProperty("HasValue").GetValue($inst)
if (-not $prop) {
    Write-Host "FAIL: HasValue should be true"
    exit 1
}
Write-Host "PASS"
exit 0
