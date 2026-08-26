# vybe-test: powershell/collections_concurrent_dictionary/concurrent_dictionary_tryremove_missing_key_returns_false
$cd = [System.Collections.Concurrent.ConcurrentDictionary[string, int]]::new()
[int]$outVal = 0
$ok = $cd.TryRemove("missing", [ref]$outVal)
if ($ok) { Write-Host "FAIL: TryRemove should return false on missing"; exit 1 }
Write-Host "PASS"; exit 0
