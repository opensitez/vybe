# vybe-test: powershell/collections_sorted_dictionary/automatic_key_sorting_order
$sd = [System.Collections.Generic.SortedDictionary[string, int]]::new()
$sd.Add("zebra", 100); $sd.Add("apple", 200); $sd.Add("mango", 300)
$keys = @($sd.Keys)
if ($keys[0] -ne "apple" -or $keys[1] -ne "mango" -or $keys[2] -ne "zebra") {
    Write-Host "FAIL: SortedDictionary key ordering failed"
    exit 1
}
Write-Host "PASS"
exit 0
