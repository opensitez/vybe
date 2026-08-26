# vybe-test: powershell/collections_immutable_arrays/immutable_array_removeat_index
$a1 = [System.Collections.Immutable.ImmutableArray]::Create([string[]]@("first", "second", "third"))
$a2 = $a1.RemoveAt(1)
if ($a2.Length -ne 2 -or $a2[1] -ne "third") { Write-Host "FAIL: RemoveAt failed"; exit 1 }
Write-Host "PASS"; exit 0
