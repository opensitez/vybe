# vybe-test: powershell/collections_linked_list/linked_list_removefirst_and_removelast
$ll = [System.Collections.Generic.LinkedList[int]]::new([int[]]@(1, 2, 3))
$ll.RemoveFirst()
$ll.RemoveLast()
if ($ll.Count -ne 1 -or $ll.First.Value -ne 2) { Write-Host "FAIL: RemoveFirst/RemoveLast failed"; exit 1 }
Write-Host "PASS"; exit 0
