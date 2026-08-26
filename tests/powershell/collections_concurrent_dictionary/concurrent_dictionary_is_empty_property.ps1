# vybe-test: powershell/collections_concurrent_dictionary/concurrent_dictionary_is_empty_property
$cd = [System.Collections.Concurrent.ConcurrentDictionary[string, string]]::new()
if (-not $cd.IsEmpty) { Write-Host "FAIL: Initial IsEmpty should be true"; exit 1 }
$cd.TryAdd("k", "v")
if ($cd.IsEmpty) { Write-Host "FAIL: IsEmpty after add should be false"; exit 1 }
Write-Host "PASS"; exit 0
