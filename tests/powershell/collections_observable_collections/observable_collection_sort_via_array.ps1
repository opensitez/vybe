# vybe-test: powershell/collections_observable_collections/observable_collection_sort_via_array
$oc = [System.Collections.ObjectModel.ObservableCollection[int]]::new([int[]]@(5, 2, 8, 1))
$sorted = @($oc | Sort-Object)
if ($sorted[0] -ne 1 -or $sorted[3] -ne 8) { Write-Host "FAIL: Sort-Object failed"; exit 1 }
Write-Host "PASS"; exit 0
