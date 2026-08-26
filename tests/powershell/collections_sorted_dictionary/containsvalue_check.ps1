# vybe-test: powershell/collections_sorted_dictionary/containsvalue_check
$sd = [System.Collections.Generic.SortedDictionary[string, int]]::new()
$sd.Add("k1", 42)
if (-not $sd.ContainsValue(42) -or $sd.ContainsValue(99)) {
    Write-Host "FAIL: ContainsValue check failed"
    exit 1
}
Write-Host "PASS"
exit 0
