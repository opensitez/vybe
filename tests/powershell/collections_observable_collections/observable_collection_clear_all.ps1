# vybe-test: powershell/collections_observable_collections/observable_collection_clear_all
$oc = [System.Collections.ObjectModel.ObservableCollection[int]]::new()
$oc.Add(1); $oc.Add(2)
$oc.Clear()
if ($oc.Count -ne 0) { Write-Host "FAIL: Clear failed"; exit 1 }
Write-Host "PASS"; exit 0
