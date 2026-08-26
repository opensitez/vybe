# vybe-test: powershell/collections_keyvaluepair_struct/kvp_deconstruct_into_tuple_variables
$kvp = [System.Collections.Generic.KeyValuePair[string, int]]::new("total", 100)
$k = $kvp.Key
$v = $kvp.Value
if ($k -ne "total" -or $v -ne 100) {
    Write-Host "FAIL: KeyValuePair property access failed"
    exit 1
}
Write-Host "PASS"
exit 0
