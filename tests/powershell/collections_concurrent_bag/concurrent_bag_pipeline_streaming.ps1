# vybe-test: powershell/collections_concurrent_bag/concurrent_bag_pipeline_streaming
$bag = [System.Collections.Concurrent.ConcurrentBag[int]]::new([int[]]@(1..5))
$collected = @($bag | ForEach-Object { $_ * 10 })
if ($collected.Length -ne 5) { Write-Host "FAIL: Pipeline streaming failed"; exit 1 }
Write-Host "PASS"; exit 0
