# vybe-test: powershell/collections_linked_list/linked_list_measure_object_sum
$ll = [System.Collections.Generic.LinkedList[int]]::new([int[]]@(1, 2, 3, 4))
$m = $ll | Measure-Object -Sum
if ($m.Sum -ne 10) { Write-Host "FAIL: Measure-Object failed"; exit 1 }
Write-Host "PASS"; exit 0
