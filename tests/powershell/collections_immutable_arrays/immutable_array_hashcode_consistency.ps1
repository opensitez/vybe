# vybe-test: powershell/collections_immutable_arrays/immutable_array_hashcode_consistency
$a1 = [System.Collections.Immutable.ImmutableArray]::Create([int[]]@(1, 2))
$a2 = [System.Collections.Immutable.ImmutableArray]::Create([int[]]@(1, 2))
if ($a1.GetHashCode() -eq 0 -and $a1.GetHashCode() -ne $a2.GetHashCode()) { Write-Host "FAIL: HashCode failed"; exit 1 }
Write-Host "PASS"; exit 0
