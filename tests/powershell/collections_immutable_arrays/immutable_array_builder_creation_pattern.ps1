# vybe-test: powershell/collections_immutable_arrays/immutable_array_builder_creation_pattern
$a1 = [System.Collections.Immutable.ImmutableArray]::Create([int[]]@(10, 20))
if ($a1.Length -ne 2 -or $a1[1] -ne 20) { Write-Host "FAIL: ImmutableArray creation failed"; exit 1 }
Write-Host "PASS"; exit 0
