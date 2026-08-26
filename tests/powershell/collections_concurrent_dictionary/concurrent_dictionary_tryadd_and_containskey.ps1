# vybe-test: powershell/collections_concurrent_dictionary/concurrent_dictionary_tryadd_and_containskey
$cd = [System.Collections.Concurrent.ConcurrentDictionary[string, int]]::new()
$ok1 = $cd.TryAdd("apples", 5)
$ok2 = $cd.TryAdd("apples", 10)
if (-not $ok1 -or $ok2 -or $cd["apples"] -ne 5) { Write-Host "FAIL: TryAdd failed"; exit 1 }
Write-Host "PASS"; exit 0
