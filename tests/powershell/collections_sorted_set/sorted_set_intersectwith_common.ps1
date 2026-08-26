# vybe-test: powershell/collections_sorted_set/sorted_set_intersectwith_common
$s1 = [System.Collections.Generic.SortedSet[int]]::new([int[]]@(1, 2, 3, 4))
$s2 = [System.Collections.Generic.SortedSet[int]]::new([int[]]@(3, 4, 5, 6))
$s1.IntersectWith($s2)
if ($s1.Count -ne 2 -or $s1.Min -ne 3 -or $s1.Max -ne 4) { Write-Host "FAIL: IntersectWith failed"; exit 1 }
Write-Host "PASS"; exit 0
