# vybe-test: powershell/collections_sorted_set/sorted_set_issupersetof_check
$sub = [System.Collections.Generic.SortedSet[int]]::new([int[]]@(1, 2))
$sup = [System.Collections.Generic.SortedSet[int]]::new([int[]]@(1, 2, 3))
if (-not $sup.IsSupersetOf($sub) -or $sub.IsSupersetOf($sup)) { Write-Host "FAIL: IsSupersetOf failed"; exit 1 }
Write-Host "PASS"; exit 0
