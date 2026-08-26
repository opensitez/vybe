# vybe-test: powershell/collections_immutable_arrays/immutable_array_addrange_from_array
$a1 = [System.Collections.Immutable.ImmutableArray]::Create([int[]]@(1, 2))
$a2 = $a1.AddRange([int[]]@(3, 4))
if ($a2.Length -ne 4 -or $a2[3] -ne 4) { Write-Host "FAIL: AddRange failed"; exit 1 }
Write-Host "PASS"; exit 0
