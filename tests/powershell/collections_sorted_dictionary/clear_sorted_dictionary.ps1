# vybe-test: powershell/collections_sorted_dictionary/clear_sorted_dictionary
$sd = [System.Collections.Generic.SortedDictionary[int, int]]::new()
$sd.Add(1, 1); $sd.Add(2, 2)
$sd.Clear()
if ($sd.Count -ne 0) {
    Write-Host "FAIL: Clear failed"
    exit 1
}
Write-Host "PASS"
exit 0
