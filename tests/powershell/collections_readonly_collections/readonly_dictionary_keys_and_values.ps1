# vybe-test: powershell/collections_readonly_collections/readonly_dictionary_keys_and_values
$d = [System.Collections.Generic.Dictionary[string, int]]::new()
$d.Add("a", 1); $d.Add("b", 2)
$rod = [System.Collections.ObjectModel.ReadOnlyDictionary[string, int]]::new($d)
$keys = @($rod.Keys)
$vals = @($rod.Values)
if ($keys.Length -ne 2 -or $vals.Length -ne 2) { Write-Host "FAIL: ReadOnlyDictionary Keys/Values failed"; exit 1 }
Write-Host "PASS"; exit 0
