# vybe-test: powershell/collections_sorted_dictionary/values_enumeration_sorted_by_key
$sd = [System.Collections.Generic.SortedDictionary[string, int]]::new()
$sd.Add("z", 26); $sd.Add("a", 1); $sd.Add("m", 13)
$vals = @($sd.Values)
if ($vals[0] -ne 1 -or $vals[1] -ne 13 -or $vals[2] -ne 26) {
    Write-Host "FAIL: Values enumeration sorted by key failed"
    exit 1
}
Write-Host "PASS"
exit 0
