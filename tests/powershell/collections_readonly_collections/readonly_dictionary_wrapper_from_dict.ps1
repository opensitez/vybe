# vybe-test: powershell/collections_readonly_collections/readonly_dictionary_wrapper_from_dict
$d = [System.Collections.Generic.Dictionary[string, int]]::new()
$d.Add("score", 100)
$rod = [System.Collections.ObjectModel.ReadOnlyDictionary[string, int]]::new($d)
if ($rod["score"] -ne 100 -or -not $rod.ContainsKey("score")) { Write-Host "FAIL: ReadOnlyDictionary wrapper failed"; exit 1 }
Write-Host "PASS"; exit 0
