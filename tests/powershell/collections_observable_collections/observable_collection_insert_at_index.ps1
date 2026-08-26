# vybe-test: powershell/collections_observable_collections/observable_collection_insert_at_index
$oc = [System.Collections.ObjectModel.ObservableCollection[int]]::new()
$oc.Add(1); $oc.Add(3)
$oc.Insert(1, 2)
if ($oc[1] -ne 2 -or $oc.Count -ne 3) { Write-Host "FAIL: Insert failed"; exit 1 }
Write-Host "PASS"; exit 0
