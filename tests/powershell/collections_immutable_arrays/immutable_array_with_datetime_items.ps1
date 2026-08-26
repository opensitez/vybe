# vybe-test: powershell/collections_immutable_arrays/immutable_array_with_datetime_items
$dt = [datetime]::UtcNow
$arr = [System.Collections.Immutable.ImmutableArray]::Create([datetime[]]@($dt))
if ($arr[0] -ne $dt) { Write-Host "FAIL: DateTime array failed"; exit 1 }
Write-Host "PASS"; exit 0
