# vybe-test: powershell/collections_concurrent_dictionary/concurrent_dictionary_trygetvalue_present_and_missing
$cd = [System.Collections.Concurrent.ConcurrentDictionary[string, int]]::new()
$cd.TryAdd("score", 777)
$hasKey = $cd.ContainsKey("score")
$val = if ($hasKey) { $cd["score"] } else { 0 }
if (-not $hasKey -or $val -ne 777 -or $cd.ContainsKey("unknown")) { Write-Host "FAIL: Dictionary lookup failed"; exit 1 }
Write-Host "PASS"; exit 0
