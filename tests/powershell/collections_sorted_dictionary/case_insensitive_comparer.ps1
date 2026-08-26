# vybe-test: powershell/collections_sorted_dictionary/case_insensitive_comparer
$sd = [System.Collections.Generic.SortedDictionary[string, int]]::new([System.StringComparer]::OrdinalIgnoreCase)
$sd.Add("AAA", 1)
if ($sd["aaa"] -ne 1) {
    Write-Host "FAIL: Case-insensitive SortedDictionary lookup failed"
    exit 1
}
Write-Host "PASS"
exit 0
