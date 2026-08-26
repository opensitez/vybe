# vybe-test: powershell/collections_linked_list/linked_list_clear_all
$ll = [System.Collections.Generic.LinkedList[int]]::new([int[]]@(1, 2))
$ll.Clear()
if ($ll.Count -ne 0 -or $ll.First -ne $null) { Write-Host "FAIL: Clear failed"; exit 1 }
Write-Host "PASS"; exit 0
