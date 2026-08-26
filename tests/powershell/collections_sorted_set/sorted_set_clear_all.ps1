# vybe-test: powershell/collections_sorted_set/sorted_set_clear_all
$ss = [System.Collections.Generic.SortedSet[int]]::new([int[]]@(1, 2, 3))
$ss.Clear()
if ($ss.Count -ne 0) { Write-Host "FAIL: SortedSet Clear failed"; exit 1 }
Write-Host "PASS"; exit 0
