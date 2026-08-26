# vybe-test: powershell/collections_concurrent_dictionary/concurrent_dictionary_multithreaded_adds_with_jobs
$cd = [System.Collections.Concurrent.ConcurrentDictionary[int, int]]::new()
for ($i = 0; $i -lt 20; $i++) { $cd.TryAdd($i, $i) }
if ($cd.Count -ne 20) { Write-Host "FAIL: Multithreaded test failed"; exit 1 }
Write-Host "PASS"; exit 0
