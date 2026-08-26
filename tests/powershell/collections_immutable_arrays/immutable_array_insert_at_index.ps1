# vybe-test: powershell/collections_immutable_arrays/immutable_array_insert_at_index
$a1 = [System.Collections.Immutable.ImmutableArray]::Create([string[]]@("a", "c"))
$a2 = $a1.Insert(1, "b")
if ($a2.Length -ne 3 -or $a2[1] -ne "b") { Write-Host "FAIL: Insert failed"; exit 1 }
Write-Host "PASS"; exit 0
