# vybe-test: powershell/collections_immutable_arrays/immutable_array_remove_element
$a1 = [System.Collections.Immutable.ImmutableArray]::Create([int[]]@(10, 20, 30))
$a2 = $a1.Remove(20)
if ($a2.Length -ne 2 -or $a2.Contains(20)) { Write-Host "FAIL: Remove failed"; exit 1 }
Write-Host "PASS"; exit 0
