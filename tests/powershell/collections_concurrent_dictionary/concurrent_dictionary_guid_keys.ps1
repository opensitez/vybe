# vybe-test: powershell/collections_concurrent_dictionary/concurrent_dictionary_guid_keys
$cd = [System.Collections.Concurrent.ConcurrentDictionary[guid, string]]::new()
$g = [guid]::NewGuid()
$cd.TryAdd($g, "GuidValue")
if ($cd[$g] -ne "GuidValue") { Write-Host "FAIL: Guid key failed"; exit 1 }
Write-Host "PASS"; exit 0
