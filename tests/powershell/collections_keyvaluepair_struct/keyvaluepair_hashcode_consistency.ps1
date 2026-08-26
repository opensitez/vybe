# vybe-test: powershell/collections_keyvaluepair_struct/keyvaluepair_hashcode_consistency
$k1 = [System.Collections.Generic.KeyValuePair[string, int]]::new("key", 5)
$k2 = [System.Collections.Generic.KeyValuePair[string, int]]::new("key", 5)
if ($k1.GetHashCode() -ne $k2.GetHashCode()) {
    Write-Host "FAIL: KeyValuePair HashCode consistency failed"
    exit 1
}
Write-Host "PASS"
exit 0
