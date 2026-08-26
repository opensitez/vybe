# vybe-test: powershell/collections_concurrent_dictionary/concurrent_dictionary_with_custom_comparer
$comp = [System.StringComparer]::OrdinalIgnoreCase
$cd = [System.Collections.Concurrent.ConcurrentDictionary[string, int]]::new($comp)
$cd.TryAdd("KEY", 42)
if ($cd["key"] -ne 42) { Write-Host "FAIL: Case insensitive ConcurrentDictionary failed"; exit 1 }
Write-Host "PASS"; exit 0
