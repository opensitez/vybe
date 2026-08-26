# vybe-test: powershell/collections_concurrent_dictionary/concurrent_dictionary_count_property
$cd = [System.Collections.Concurrent.ConcurrentDictionary[int, int]]::new()
for ($i = 0; $i -lt 10; $i++) { $cd.TryAdd($i, $i * 10) }
if ($cd.Count -ne 10) { Write-Host "FAIL: Count property failed"; exit 1 }
Write-Host "PASS"; exit 0
