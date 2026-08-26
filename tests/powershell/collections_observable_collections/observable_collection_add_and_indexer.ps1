# vybe-test: powershell/collections_observable_collections/observable_collection_add_and_indexer
$oc = [System.Collections.ObjectModel.ObservableCollection[int]]::new()
$oc.Add(10); $oc.Add(20)
if ($oc.Count -ne 2 -or $oc[0] -ne 10 -or $oc[1] -ne 20) { Write-Host "FAIL: ObservableCollection Add failed"; exit 1 }
Write-Host "PASS"; exit 0
