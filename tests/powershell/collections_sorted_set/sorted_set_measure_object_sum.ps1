# vybe-test: powershell/collections_sorted_set/sorted_set_measure_object_sum
$ss = [System.Collections.Generic.SortedSet[int]]::new([int[]]@(10, 20, 30))
$m = $ss | Measure-Object -Sum
if ($m.Sum -ne 60) { Write-Host "FAIL: Measure-Object failed"; exit 1 }
Write-Host "PASS"; exit 0
