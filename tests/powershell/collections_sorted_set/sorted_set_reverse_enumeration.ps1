# vybe-test: powershell/collections_sorted_set/sorted_set_reverse_enumeration
$ss = [System.Collections.Generic.SortedSet[int]]::new([int[]]@(1, 2, 3))
$rev = @($ss.Reverse())
if ($rev[0] -ne 3 -or $rev[2] -ne 1) { Write-Host "FAIL: Reverse enumeration failed"; exit 1 }
Write-Host "PASS"; exit 0
