# vybe-test: powershell/collections_sorted_dictionary/reference_equality_sorted_dict
$sd1 = [System.Collections.Generic.SortedDictionary[string, string]]::new()
$sd2 = $sd1
if ($sd1 -ne $sd2) {
    Write-Host "FAIL: Reference equality failed"
    exit 1
}
Write-Host "PASS"
exit 0
