# vybe-test: powershell/collections_immutable_arrays/immutable_array_pipeline_filter
$arr = [System.Collections.Immutable.ImmutableArray]::Create([int[]]@(1..6))
$evens = @($arr | Where-Object { $_ % 2 -eq 0 })
if ($evens.Length -ne 3) { Write-Host "FAIL: Pipeline filter failed"; exit 1 }
Write-Host "PASS"; exit 0
