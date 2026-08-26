# vybe-test: powershell/collections_linked_list/linked_list_node_value_mutation
$ll = [System.Collections.Generic.LinkedList[string]]::new([string[]]@("old"))
$ll.First.Value = "new"
if ($ll.First.Value -ne "new") { Write-Host "FAIL: Node value mutation failed"; exit 1 }
Write-Host "PASS"; exit 0
