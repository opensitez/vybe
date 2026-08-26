# vybe-test: powershell/collections_keyvaluepair_struct/inequality_different_values
$k1 = [System.Collections.Generic.KeyValuePair[string, int]]::new("a", 1)
$k2 = [System.Collections.Generic.KeyValuePair[string, int]]::new("a", 2)
if ($k1 -eq $k2) {
    Write-Host "FAIL: Different values must compare unequal"
    exit 1
}
Write-Host "PASS"
exit 0
