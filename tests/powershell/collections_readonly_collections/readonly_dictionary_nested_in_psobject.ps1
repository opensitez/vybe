# vybe-test: powershell/collections_readonly_collections/readonly_dictionary_nested_in_psobject
$d = [System.Collections.Generic.Dictionary[string, int]]::new()
$d.Add("val", 99)
$obj = [pscustomobject]@{ ReadOnlyMap = [System.Collections.ObjectModel.ReadOnlyDictionary[string, int]]::new($d) }
if ($obj.ReadOnlyMap["val"] -ne 99) { Write-Host "FAIL: Nested ReadOnlyDictionary in PSCustomObject failed"; exit 1 }
Write-Host "PASS"; exit 0
