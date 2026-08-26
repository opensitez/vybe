# vybe-test: powershell/collections_concurrent_dictionary/concurrent_dictionary_complex_object_values
$cd = [System.Collections.Concurrent.ConcurrentDictionary[string, pscustomobject]]::new()
$cd.TryAdd("obj", [pscustomobject]@{ Id = 1; Name = "Test" })
if ($cd["obj"].Name -ne "Test") { Write-Host "FAIL: Complex object value failed"; exit 1 }
Write-Host "PASS"; exit 0
