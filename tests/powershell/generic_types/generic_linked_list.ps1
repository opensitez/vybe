# vybe-test: powershell/generic_types/generic_linked_list
$ll = [System.Collections.Generic.LinkedList[string]]::new()
$node = $ll.AddFirst("Head")
[void]$ll.AddAfter($node, "Tail")
if ($ll.Last.Value -ne "Tail") {
    Write-Host "FAIL: LinkedList Last.Value expected Tail, got $($ll.Last.Value)"
    exit 1
}
Write-Host "PASS"
exit 0
