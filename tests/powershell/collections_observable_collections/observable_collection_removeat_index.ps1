# vybe-test: powershell/collections_observable_collections/observable_collection_removeat_index
$oc = [System.Collections.ObjectModel.ObservableCollection[int]]::new()
$oc.Add(10); $oc.Add(20); $oc.Add(30)
$oc.RemoveAt(1)
if ($oc.Count -ne 2 -or $oc[1] -ne 30) { Write-Host "FAIL: RemoveAt failed"; exit 1 }
Write-Host "PASS"; exit 0
