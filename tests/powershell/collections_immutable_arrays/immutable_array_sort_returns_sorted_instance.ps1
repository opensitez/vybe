# vybe-test: powershell/collections_immutable_arrays/immutable_array_sort_returns_sorted_instance
$a1 = [System.Collections.Immutable.ImmutableArray]::Create([int[]]@(5, 2, 8, 1))
$a2 = $a1.Sort()
if ($a2[0] -ne 1 -or $a2[3] -ne 8 -or $a1[0] -ne 5) { Write-Host "FAIL: Sort immutability failed"; exit 1 }
Write-Host "PASS"; exit 0
