# vybe-test: powershell/collections_linked_list/linked_list_find_node
$ll = [System.Collections.Generic.LinkedList[string]]::new([string[]]@("a", "target", "c"))
$node = $ll.Find("target")
if ($node -eq $null -or $node.Value -ne "target" -or $node.Next.Value -ne "c") { Write-Host "FAIL: Find node failed"; exit 1 }
Write-Host "PASS"; exit 0
