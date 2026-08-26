# vybe-test: powershell/collections_keyvaluepair_struct/kvp_in_array_filter
$k1 = [System.Collections.Generic.KeyValuePair[string, int]]::new("a", 10)
$k2 = [System.Collections.Generic.KeyValuePair[string, int]]::new("b", 20)
$arr = @($k1, $k2)
$filtered = @($arr | Where-Object { $_.Value -gt 15 })
if ($filtered.Length -ne 1 -or $filtered[0].Key -ne "b") {
    Write-Host "FAIL: KeyValuePair array filter failed"
    exit 1
}
Write-Host "PASS"
exit 0
