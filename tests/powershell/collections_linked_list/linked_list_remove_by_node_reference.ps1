# vybe-test: powershell/collections_linked_list/linked_list_remove_by_node_reference
$ll = [System.Collections.Generic.LinkedList[int]]::new([int[]]@(1, 2, 3))
$node = $ll.Find(2)
$ll.Remove($node)
if ($ll.Count -ne 2 -or $ll.Contains(2)) { Write-Host "FAIL: Remove by node reference failed"; exit 1 }
Write-Host "PASS"; exit 0
