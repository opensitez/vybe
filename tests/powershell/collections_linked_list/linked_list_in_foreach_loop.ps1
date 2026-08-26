# vybe-test: powershell/collections_linked_list/linked_list_in_foreach_loop
$ll = [System.Collections.Generic.LinkedList[int]]::new([int[]]@(10, 20, 30))
$sum = 0
foreach ($item in $ll) { $sum += $item }
if ($sum -ne 60) { Write-Host "FAIL: foreach on LinkedList failed"; exit 1 }
Write-Host "PASS"; exit 0
