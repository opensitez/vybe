# vybe-test: powershell/collections_concurrent_dictionary/concurrent_dictionary_indexer_read_and_write
$cd = [System.Collections.Concurrent.ConcurrentDictionary[string, string]]::new()
$cd["greeting"] = "hello"
$cd["greeting"] = "world"
if ($cd["greeting"] -ne "world") { Write-Host "FAIL: Indexer read/write failed"; exit 1 }
Write-Host "PASS"; exit 0
