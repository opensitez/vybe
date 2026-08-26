# vybe-test: powershell/collections_observable_collections/observable_collection_copyto_array
$oc = [System.Collections.ObjectModel.ObservableCollection[int]]::new([int[]]@(1, 2))
[int[]]$arr = [int[]]::new(2)
$oc.CopyTo($arr, 0)
if ($arr[0] -ne 1 -or $arr[1] -ne 2) { Write-Host "FAIL: CopyTo failed"; exit 1 }
Write-Host "PASS"; exit 0
