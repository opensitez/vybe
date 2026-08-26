# vybe-test: powershell/collections_concurrent_dictionary/concurrent_dictionary_null_value_handling
$cd = [System.Collections.Concurrent.ConcurrentDictionary[string, string]]::new()
$cd.TryAdd("validKey", "validValue")
$hasVal = $cd.ContainsKey("validKey")
$noVal = $cd.ContainsKey("nonExistentKey")
if (-not $hasVal -or $noVal) {
    Write-Host "FAIL: Key check in ConcurrentDictionary failed"
    exit 1
}
Write-Host "PASS"
exit 0
