# vybe-test: powershell/collections_sorted_set/sorted_set_issubsetof_check
$sub = [System.Collections.Generic.SortedSet[int]]::new([int[]]@(1, 2))
$sup = [System.Collections.Generic.SortedSet[int]]::new([int[]]@(1, 2, 3))
if (-not $sub.IsSubsetOf($sup) -or $sup.IsSubsetOf($sub)) { Write-Host "FAIL: IsSubsetOf failed"; exit 1 }
Write-Host "PASS"; exit 0
