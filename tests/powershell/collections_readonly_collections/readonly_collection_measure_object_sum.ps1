# vybe-test: powershell/collections_readonly_collections/readonly_collection_measure_object_sum
$list = [System.Collections.Generic.List[int]]::new([int[]]@(10, 20, 30))
$roc = [System.Collections.ObjectModel.ReadOnlyCollection[int]]::new($list)
$m = $roc | Measure-Object -Sum
if ($m.Sum -ne 60) { Write-Host "FAIL: Measure-Object failed"; exit 1 }
Write-Host "PASS"; exit 0
