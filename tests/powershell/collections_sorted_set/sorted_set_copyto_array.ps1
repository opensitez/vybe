# vybe-test: powershell/collections_sorted_set/sorted_set_copyto_array
$ss = [System.Collections.Generic.SortedSet[int]]::new([int[]]@(30, 10, 20))
[int[]]$arr = [int[]]::new(3)
$ss.CopyTo($arr)
if ($arr[0] -ne 10 -or $arr[1] -ne 20 -or $arr[2] -ne 30) { Write-Host "FAIL: CopyTo failed"; exit 1 }
Write-Host "PASS"; exit 0
