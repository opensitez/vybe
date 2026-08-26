# vybe-test: powershell/collections_immutable_arrays/immutable_array_binarysearch_element
$arr = [System.Collections.Immutable.ImmutableArray]::Create([int[]]@(10, 20, 30, 40, 50))
$raw = @($arr)
$idx = $raw.IndexOf(30)
if ($idx -ne 2) { Write-Host "FAIL: IndexOf search failed"; exit 1 }
Write-Host "PASS"; exit 0
