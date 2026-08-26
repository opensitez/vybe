# vybe-test: powershell/collections_linked_list/linked_list_findlast_node
$ll = [System.Collections.Generic.LinkedList[int]]::new([int[]]@(1, 2, 3, 2, 4))
$node = $ll.FindLast(2)
if ($node -eq $null -or $node.Next.Value -ne 4) { Write-Host "FAIL: FindLast failed"; exit 1 }
Write-Host "PASS"; exit 0
