# vybe-test: powershell/collections_observable_collections/observable_collection_indexer_set
$oc = [System.Collections.ObjectModel.ObservableCollection[string]]::new()
$oc.Add("old")
$oc[0] = "new"
if ($oc[0] -ne "new") { Write-Host "FAIL: Indexer set failed"; exit 1 }
Write-Host "PASS"; exit 0
