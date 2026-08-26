# vybe-test: powershell/collections_linked_list/linked_list_with_complex_objects
$ll = [System.Collections.Generic.LinkedList[pscustomobject]]::new()
$null = $ll.AddLast([pscustomobject]@{ Code = "ALPHA" })
if ($ll.First.Value.Code -ne "ALPHA") { Write-Host "FAIL: Complex object in LinkedList failed"; exit 1 }
Write-Host "PASS"; exit 0
