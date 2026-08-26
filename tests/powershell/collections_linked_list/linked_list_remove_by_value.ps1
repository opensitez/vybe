# vybe-test: powershell/collections_linked_list/linked_list_remove_by_value
$ll = [System.Collections.Generic.LinkedList[string]]::new([string[]]@("x", "y", "z"))
$rem = $ll.Remove("y")
if (-not $rem -or $ll.Count -ne 2 -or $ll.Contains("y")) { Write-Host "FAIL: Remove by value failed"; exit 1 }
Write-Host "PASS"; exit 0
