# vybe-test: powershell/collections_immutable_arrays/immutable_array_is_empty_property
$empty = [System.Collections.Immutable.ImmutableArray[int]]::Empty
if (-not $empty.IsEmpty -or $empty.Length -ne 0) { Write-Host "FAIL: Empty failed"; exit 1 }
Write-Host "PASS"; exit 0
