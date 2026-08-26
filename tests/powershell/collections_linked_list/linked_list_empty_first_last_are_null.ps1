# vybe-test: powershell/collections_linked_list/linked_list_empty_first_last_are_null
$ll = [System.Collections.Generic.LinkedList[int]]::new()
if ($ll.First -ne $null -or $ll.Last -ne $null) { Write-Host "FAIL: Empty first/last should be null"; exit 1 }
Write-Host "PASS"; exit 0
