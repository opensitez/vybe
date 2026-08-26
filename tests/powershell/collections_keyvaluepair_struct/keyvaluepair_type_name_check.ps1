# vybe-test: powershell/collections_keyvaluepair_struct/keyvaluepair_type_name_check
$kvp = [System.Collections.Generic.KeyValuePair[string, int]]::new("test", 123)
if (-not ($kvp.GetType().Name.StartsWith("KeyValuePair"))) {
    Write-Host "FAIL: Type name check failed"
    exit 1
}
Write-Host "PASS"
exit 0
