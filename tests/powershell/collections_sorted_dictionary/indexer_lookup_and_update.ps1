# vybe-test: powershell/collections_sorted_dictionary/indexer_lookup_and_update
$sd = [System.Collections.Generic.SortedDictionary[string, string]]::new()
$sd["host"] = "prod"
$sd["host"] = "staging"
if ($sd["host"] -ne "staging" -or $sd.Count -ne 1) {
    Write-Host "FAIL: Indexer update failed"
    exit 1
}
Write-Host "PASS"
exit 0
