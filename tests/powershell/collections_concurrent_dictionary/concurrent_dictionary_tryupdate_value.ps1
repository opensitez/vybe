# vybe-test: powershell/collections_concurrent_dictionary/concurrent_dictionary_tryupdate_value
$cd = [System.Collections.Concurrent.ConcurrentDictionary[string, int]]::new()
$cd.TryAdd("level", 1)
$ok = $cd.TryUpdate("level", 2, 1)
$fail = $cd.TryUpdate("level", 3, 1)
if (-not $ok -or $fail -or $cd["level"] -ne 2) { Write-Host "FAIL: TryUpdate failed"; exit 1 }
Write-Host "PASS"; exit 0
