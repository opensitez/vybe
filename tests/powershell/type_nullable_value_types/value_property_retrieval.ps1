# vybe-test: powershell/type_nullable_value_types/value_property_retrieval
$t = [type]"System.Nullable[int]"
$inst = [System.Activator]::CreateInstance($t, @(100))
$val = $t.GetProperty("Value").GetValue($inst)
if ($val -ne 100) {
    Write-Host "FAIL: Value retrieval failed"
    exit 1
}
Write-Host "PASS"
exit 0
