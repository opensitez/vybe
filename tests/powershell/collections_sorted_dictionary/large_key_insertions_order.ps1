# vybe-test: powershell/collections_sorted_dictionary/large_key_insertions_order
$sd = [System.Collections.Generic.SortedDictionary[int, int]]::new()
for ($i = 50; $i -ge 1; $i--) {
    $sd.Add($i, $i * 2)
}
$keys = @($sd.Keys)
if ($keys[0] -ne 1 -or $keys[49] -ne 50 -or $sd[25] -ne 50) {
    Write-Host "FAIL: Large key insertions sorted check failed"
    exit 1
}
Write-Host "PASS"
exit 0
