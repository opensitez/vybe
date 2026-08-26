# vybe-test: powershell/collections_concurrent_dictionary/concurrent_dictionary_getoradd_factory
$cd = [System.Collections.Concurrent.ConcurrentDictionary[string, int]]::new()
$val = $cd.GetOrAdd("score", 100)
$val2 = $cd.GetOrAdd("score", 200)
if ($val -ne 100 -or $val2 -ne 100) { Write-Host "FAIL: GetOrAdd failed"; exit 1 }
Write-Host "PASS"; exit 0
