# vybe-test: powershell/collections_concurrent_dictionary/concurrent_dictionary_getoradd_existing_does_not_call_factory
$cd = [System.Collections.Concurrent.ConcurrentDictionary[string, int]]::new()
$cd.TryAdd("k", 50)
$val = $cd.GetOrAdd("k", 999)
if ($val -ne 50) { Write-Host "FAIL: GetOrAdd existing key should return original value"; exit 1 }
Write-Host "PASS"; exit 0
