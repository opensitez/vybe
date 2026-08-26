# vybe-test: powershell/collections_concurrent_dictionary/concurrent_dictionary_toarray_keyvaluepairs
$cd = [System.Collections.Concurrent.ConcurrentDictionary[string, int]]::new()
$cd.TryAdd("x", 10); $cd.TryAdd("y", 20)
$arr = $cd.ToArray()
if ($arr.Length -ne 2) { Write-Host "FAIL: ToArray failed"; exit 1 }
Write-Host "PASS"; exit 0
