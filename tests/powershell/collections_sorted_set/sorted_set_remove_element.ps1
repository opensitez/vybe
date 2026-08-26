# vybe-test: powershell/collections_sorted_set/sorted_set_remove_element
$ss = [System.Collections.Generic.SortedSet[string]]::new([string[]]@("a", "b", "c"))
$rem = $ss.Remove("b")
if (-not $rem -or $ss.Count -ne 2 -or $ss.Contains("b")) { Write-Host "FAIL: SortedSet Remove failed"; exit 1 }
Write-Host "PASS"; exit 0
