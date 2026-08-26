# vybe-test: powershell/collections_sorted_set/sorted_set_min_and_max_properties
$ss = [System.Collections.Generic.SortedSet[int]]::new([int[]]@(50, 10, 90, 30))
if ($ss.Min -ne 10 -or $ss.Max -ne 90) { Write-Host "FAIL: SortedSet Min/Max failed"; exit 1 }
Write-Host "PASS"; exit 0
