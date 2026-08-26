# vybe-test: powershell/collections_sorted_set/sorted_set_symmetricexceptwith_xor
$s1 = [System.Collections.Generic.SortedSet[int]]::new([int[]]@(1, 2, 3))
$s2 = [System.Collections.Generic.SortedSet[int]]::new([int[]]@(2, 3, 4))
$s1.SymmetricExceptWith($s2)
if ($s1.Count -ne 2 -or -not $s1.Contains(1) -or -not $s1.Contains(4)) { Write-Host "FAIL: SymmetricExceptWith failed"; exit 1 }
Write-Host "PASS"; exit 0
