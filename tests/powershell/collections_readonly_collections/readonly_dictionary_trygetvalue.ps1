# vybe-test: powershell/collections_readonly_collections/readonly_dictionary_trygetvalue
$d = [System.Collections.Generic.Dictionary[string, int]]::new()
$d.Add("k", 42)
$rod = [System.Collections.ObjectModel.ReadOnlyDictionary[string, int]]::new($d)
[int]$outVal = 0
$ok = $rod.TryGetValue("k", [ref]$outVal)
if (-not $ok -or $outVal -ne 42) { Write-Host "FAIL: ReadOnlyDictionary TryGetValue failed"; exit 1 }
Write-Host "PASS"; exit 0
