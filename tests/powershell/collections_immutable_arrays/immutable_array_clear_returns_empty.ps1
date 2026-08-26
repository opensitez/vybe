# vybe-test: powershell/collections_immutable_arrays/immutable_array_clear_returns_empty
$a1 = [System.Collections.Immutable.ImmutableArray]::Create([int[]]@(1, 2, 3))
$a2 = $a1.Clear()
if ($a2.Length -ne 0 -or -not $a2.IsEmpty) { Write-Host "FAIL: Clear failed"; exit 1 }
Write-Host "PASS"; exit 0
