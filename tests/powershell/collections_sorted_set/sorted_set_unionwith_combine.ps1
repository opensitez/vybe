# vybe-test: powershell/collections_sorted_set/sorted_set_unionwith_combine
$s1 = [System.Collections.Generic.SortedSet[int]]::new([int[]]@(1, 2, 3))
$s2 = [System.Collections.Generic.SortedSet[int]]::new([int[]]@(3, 4, 5))
$s1.UnionWith($s2)
if ($s1.Count -ne 5 -or $s1.Max -ne 5) { Write-Host "FAIL: UnionWith failed"; exit 1 }
Write-Host "PASS"; exit 0
