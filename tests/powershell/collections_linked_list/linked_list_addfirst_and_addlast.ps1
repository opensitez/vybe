# vybe-test: powershell/collections_linked_list/linked_list_addfirst_and_addlast
$ll = [System.Collections.Generic.LinkedList[string]]::new()
$null = $ll.AddLast("Middle")
$null = $ll.AddFirst("First")
$null = $ll.AddLast("Last")
if ($ll.First.Value -ne "First" -or $ll.Last.Value -ne "Last" -or $ll.Count -ne 3) { Write-Host "FAIL: AddFirst/AddLast failed"; exit 1 }
Write-Host "PASS"; exit 0
