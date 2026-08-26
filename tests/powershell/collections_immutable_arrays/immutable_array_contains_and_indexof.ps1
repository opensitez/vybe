# vybe-test: powershell/collections_immutable_arrays/immutable_array_contains_and_indexof
$arr = [System.Collections.Immutable.ImmutableArray]::Create([string[]]@("x", "y", "z"))
if (-not $arr.Contains("y") -or $arr.IndexOf("z") -ne 2) { Write-Host "FAIL: Contains/IndexOf failed"; exit 1 }
Write-Host "PASS"; exit 0
