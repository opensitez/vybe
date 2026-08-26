# vybe-test: powershell/collections_observable_collections/observable_collection_construct_from_enumerable
$oc = [System.Collections.ObjectModel.ObservableCollection[int]]::new([int[]]@(1, 2, 3, 4))
if ($oc.Count -ne 4 -or $oc[3] -ne 4) { Write-Host "FAIL: Constructor from enumerable failed"; exit 1 }
Write-Host "PASS"; exit 0
