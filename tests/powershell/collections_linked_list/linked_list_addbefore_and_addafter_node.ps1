# vybe-test: powershell/collections_linked_list/linked_list_addbefore_and_addafter_node
$ll = [System.Collections.Generic.LinkedList[int]]::new()
$node2 = $ll.AddLast(20)
$node1 = $ll.AddBefore($node2, 10)
$node3 = $ll.AddAfter($node2, 30)
if ($ll.First.Value -ne 10 -or $ll.Last.Value -ne 30) { Write-Host "FAIL: AddBefore/AddAfter failed"; exit 1 }
Write-Host "PASS"; exit 0
