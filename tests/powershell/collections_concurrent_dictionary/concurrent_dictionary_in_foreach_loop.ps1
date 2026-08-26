# vybe-test: powershell/collections_concurrent_dictionary/concurrent_dictionary_in_foreach_loop
$cd = [System.Collections.Concurrent.ConcurrentDictionary[string, int]]::new()
$cd.TryAdd("a", 1); $cd.TryAdd("b", 2)
$sum = 0
foreach ($k in $cd.Keys) { $sum += $cd[$k] }
if ($sum -ne 3) { Write-Host "FAIL: foreach iteration failed"; exit 1 }
Write-Host "PASS"; exit 0
