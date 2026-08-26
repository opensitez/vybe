# vybe-test: powershell/collections_keyvaluepair_struct/equality_same_key_and_value
$k1 = [System.Collections.Generic.KeyValuePair[string, int]]::new("a", 1)
$k2 = [System.Collections.Generic.KeyValuePair[string, int]]::new("a", 1)
if ($k1 -ne $k2) {
    Write-Host "FAIL: KeyValuePairs with same key and value must be equal"
    exit 1
}
Write-Host "PASS"
exit 0
