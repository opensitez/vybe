# vybe-test: powershell/collections_concurrent_dictionary/concurrent_dictionary_keys_and_values_collections
$cd = [System.Collections.Concurrent.ConcurrentDictionary[string, int]]::new()
$cd.TryAdd("k1", 100); $cd.TryAdd("k2", 200)
$keys = @($cd.Keys)
$vals = @($cd.Values)
if ($keys.Length -ne 2 -or $vals.Length -ne 2) { Write-Host "FAIL: Keys/Values failed"; exit 1 }
Write-Host "PASS"; exit 0
