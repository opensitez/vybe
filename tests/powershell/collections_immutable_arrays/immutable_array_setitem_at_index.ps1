# vybe-test: powershell/collections_immutable_arrays/immutable_array_setitem_at_index
$a1 = [System.Collections.Immutable.ImmutableArray]::Create([string[]]@("a", "b"))
$a2 = $a1.SetItem(0, "z")
if ($a1[0] -ne "a" -or $a2[0] -ne "z") { Write-Host "FAIL: SetItem failed"; exit 1 }
Write-Host "PASS"; exit 0
