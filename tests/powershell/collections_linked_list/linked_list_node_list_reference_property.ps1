# vybe-test: powershell/collections_linked_list/linked_list_node_list_reference_property
$ll = [System.Collections.Generic.LinkedList[int]]::new([int[]]@(100))
$node = $ll.First
if ($node.List -eq $null -or $node.Value -ne 100) { Write-Host "FAIL: Node.List reference failed"; exit 1 }
Write-Host "PASS"; exit 0
