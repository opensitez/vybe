# vybe-test: powershell/collections_concurrent_dictionary/concurrent_dictionary_clear_elements
$cd = [System.Collections.Concurrent.ConcurrentDictionary[string, int]]::new()
$cd.TryAdd("a", 1); $cd.TryAdd("b", 2)
$cd.Clear()
if ($cd.Count -ne 0 -or -not $cd.IsEmpty) { Write-Host "FAIL: ConcurrentDictionary Clear failed"; exit 1 }
Write-Host "PASS"; exit 0
