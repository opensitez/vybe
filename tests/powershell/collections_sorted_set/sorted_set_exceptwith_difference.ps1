# vybe-test: powershell/collections_sorted_set/sorted_set_exceptwith_difference
$s1 = [System.Collections.Generic.SortedSet[string]]::new([string[]]@("a", "b", "c"))
$s2 = [System.Collections.Generic.SortedSet[string]]::new([string[]]@("b"))
$s1.ExceptWith($s2)
if ($s1.Count -ne 2 -or $s1.Contains("b")) { Write-Host "FAIL: ExceptWith failed"; exit 1 }
Write-Host "PASS"; exit 0
