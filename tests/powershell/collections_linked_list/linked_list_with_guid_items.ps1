# vybe-test: powershell/collections_linked_list/linked_list_with_guid_items
$ll = [System.Collections.Generic.LinkedList[guid]]::new()
$g = [guid]::NewGuid()
$null = $ll.AddLast($g)
if ($ll.First.Value -ne $g) { Write-Host "FAIL: Guid item failed"; exit 1 }
Write-Host "PASS"; exit 0
