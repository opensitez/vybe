# vybe-test: powershell/collections_concurrent_dictionary/concurrent_dictionary_tryremove_key
$cd = [System.Collections.Concurrent.ConcurrentDictionary[string, int]]::new()
$cd.TryAdd("temp", 99)
[int]$outVal = 0
$ok = $cd.TryRemove("temp", [ref]$outVal)
if (-not $ok -or $outVal -ne 99 -or $cd.ContainsKey("temp")) { Write-Host "FAIL: TryRemove failed"; exit 1 }
Write-Host "PASS"; exit 0
