# vybe-test: powershell/collections_immutable_arrays/immutable_array_create_and_access
$arr = [System.Collections.Immutable.ImmutableArray]::Create([int[]]@(1, 2, 3))
if ($arr.Length -ne 3 -or $arr[1] -ne 2) { Write-Host "FAIL: ImmutableArray Create failed"; exit 1 }
Write-Host "PASS"; exit 0
