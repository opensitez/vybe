# vybe-test: powershell/collections_immutable_arrays/immutable_array_add_returns_new_instance
$a1 = [System.Collections.Immutable.ImmutableArray]::Create([int[]]@(1, 2))
$a2 = $a1.Add(3)
if ($a1.Length -ne 2 -or $a2.Length -ne 3 -or $a2[2] -ne 3) { Write-Host "FAIL: ImmutableArray Add immutability failed"; exit 1 }
Write-Host "PASS"; exit 0
