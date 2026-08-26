# vybe-test: powershell/collections_immutable_arrays/immutable_array_with_guid_items
$g = [guid]::NewGuid()
$arr = [System.Collections.Immutable.ImmutableArray]::Create([guid[]]@($g))
if ($arr[0] -ne $g) { Write-Host "FAIL: Guid array failed"; exit 1 }
Write-Host "PASS"; exit 0
