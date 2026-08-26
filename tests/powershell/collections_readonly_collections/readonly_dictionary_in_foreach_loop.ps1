# vybe-test: powershell/collections_readonly_collections/readonly_dictionary_in_foreach_loop
$d = [System.Collections.Generic.Dictionary[string, int]]::new()
$d.Add("x", 10); $d.Add("y", 20)
$rod = [System.Collections.ObjectModel.ReadOnlyDictionary[string, int]]::new($d)
$sum = 0
foreach ($k in $rod.Keys) { $sum += $rod[$k] }
if ($sum -ne 30) { Write-Host "FAIL: foreach on ReadOnlyDictionary failed"; exit 1 }
Write-Host "PASS"; exit 0
