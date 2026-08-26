# vybe-test: powershell/collections_linked_list/linked_list_copyto_array
$ll = [System.Collections.Generic.LinkedList[int]]::new([int[]]@(10, 20))
[int[]]$arr = [int[]]::new(2)
$ll.CopyTo($arr, 0)
if ($arr[0] -ne 10 -or $arr[1] -ne 20) { Write-Host "FAIL: CopyTo failed"; exit 1 }
Write-Host "PASS"; exit 0
