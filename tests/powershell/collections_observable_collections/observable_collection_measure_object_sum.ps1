# vybe-test: powershell/collections_observable_collections/observable_collection_measure_object_sum
$oc = [System.Collections.ObjectModel.ObservableCollection[int]]::new([int[]]@(10, 20, 30))
$m = $oc | Measure-Object -Sum
if ($m.Sum -ne 60) { Write-Host "FAIL: Measure-Object failed"; exit 1 }
Write-Host "PASS"; exit 0
