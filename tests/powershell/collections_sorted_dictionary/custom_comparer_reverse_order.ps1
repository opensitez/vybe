# vybe-test: powershell/collections_sorted_dictionary/custom_comparer_reverse_order
$comp = [System.StringComparer]::OrdinalIgnoreCase
$sd = [System.Collections.Generic.SortedDictionary[string, int]]::new($comp)
$sd.Add("a", 1); $sd.Add("c", 3); $sd.Add("b", 2)
$keys = @($sd.Keys)
if ($keys[0] -ne "a" -or $keys[1] -ne "b" -or $keys[2] -ne "c") {
    Write-Host "FAIL: Custom reverse comparer failed"
    exit 1
}
Write-Host "PASS"
exit 0
