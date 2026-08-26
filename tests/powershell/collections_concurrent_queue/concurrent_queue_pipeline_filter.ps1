# vybe-test: powershell/collections_concurrent_queue/concurrent_queue_pipeline_filter
$cq = [System.Collections.Concurrent.ConcurrentQueue[int]]::new([int[]]@(1..10))
$evens = @($cq | Where-Object { $_ % 2 -eq 0 })
if ($evens.Length -ne 5 -or $evens[0] -ne 2 -or $evens[4] -ne 10) { Write-Host "FAIL: Pipeline filter failed"; exit 1 }
Write-Host "PASS"; exit 0
