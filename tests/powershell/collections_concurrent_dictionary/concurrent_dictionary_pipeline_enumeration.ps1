# vybe-test: powershell/collections_concurrent_dictionary/concurrent_dictionary_pipeline_enumeration
$cd = [System.Collections.Concurrent.ConcurrentDictionary[string, int]]::new()
$cd.TryAdd("a", 10); $cd.TryAdd("b", 20)
$keys = @($cd.Keys | Sort-Object)
if ($keys[0] -ne "a" -or $keys[1] -ne "b") { Write-Host "FAIL: Pipeline enumeration failed"; exit 1 }
Write-Host "PASS"; exit 0
