# vybe-test: powershell/collections_linked_list/linked_list_node_next_and_previous_pointers
$ll = [System.Collections.Generic.LinkedList[int]]::new([int[]]@(10, 20, 30))
$mid = $ll.First.Next
if ($mid.Previous.Value -ne 10 -or $mid.Next.Value -ne 30) { Write-Host "FAIL: Node pointer navigation failed"; exit 1 }
Write-Host "PASS"; exit 0
