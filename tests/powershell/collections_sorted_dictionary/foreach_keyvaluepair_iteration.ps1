# vybe-test: powershell/collections_sorted_dictionary/foreach_keyvaluepair_iteration
$sd = [System.Collections.Generic.SortedDictionary[int, string]]::new()
$sd.Add(2, "two"); $sd.Add(1, "one")
$keys = @($sd.Keys)
if ($keys[0] -ne 1 -or $keys[1] -ne 2) {
    Write-Host "FAIL: SortedDictionary keys ordering failed"
    exit 1
}
Write-Host "PASS"
exit 0
