# vybe-test: powershell/collections_sorted_set/sorted_set_in_foreach_loop_sorted_order
$ss = [System.Collections.Generic.SortedSet[int]]::new([int[]]@(5, 1, 4, 2, 3))
$list = [System.Collections.Generic.List[int]]::new()
foreach ($item in $ss) { $list.Add($item) }
if ($list[0] -ne 1 -or $list[4] -ne 5) { Write-Host "FAIL: foreach sorted order failed"; exit 1 }
Write-Host "PASS"; exit 0
