# vybe-test: powershell/collections_sorted_set/sorted_set_setequals_check
$s1 = [System.Collections.Generic.SortedSet[string]]::new([string[]]@("x", "y"))
$s2 = [System.Collections.Generic.SortedSet[string]]::new([string[]]@("y", "x"))
if (-not $s1.SetEquals($s2)) { Write-Host "FAIL: SetEquals failed"; exit 1 }
Write-Host "PASS"; exit 0
