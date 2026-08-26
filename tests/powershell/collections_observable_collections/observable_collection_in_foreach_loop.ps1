# vybe-test: powershell/collections_observable_collections/observable_collection_in_foreach_loop
$oc = [System.Collections.ObjectModel.ObservableCollection[int]]::new([int[]]@(10, 20, 30))
$sum = 0
foreach ($item in $oc) { $sum += $item }
if ($sum -ne 60) { Write-Host "FAIL: foreach failed"; exit 1 }
Write-Host "PASS"; exit 0
