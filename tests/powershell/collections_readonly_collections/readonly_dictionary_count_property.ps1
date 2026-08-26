# vybe-test: powershell/collections_readonly_collections/readonly_dictionary_count_property
$d = [System.Collections.Generic.Dictionary[int, int]]::new()
for ($i = 0; $i -lt 15; $i++) { $d.Add($i, $i) }
$rod = [System.Collections.ObjectModel.ReadOnlyDictionary[int, int]]::new($d)
if ($rod.Count -ne 15) { Write-Host "FAIL: ReadOnlyDictionary Count failed"; exit 1 }
Write-Host "PASS"; exit 0
