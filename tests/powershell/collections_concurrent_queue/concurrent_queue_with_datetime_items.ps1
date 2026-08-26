# vybe-test: powershell/collections_concurrent_queue/concurrent_queue_with_datetime_items
$cq = [System.Collections.Concurrent.ConcurrentQueue[datetime]]::new()
$dt = [datetime]::UtcNow
$cq.Enqueue($dt)
[datetime]$outDt = [datetime]::MinValue
$ok = $cq.TryDequeue([ref]$outDt)
if (-not $ok -or $outDt -ne $dt) { Write-Host "FAIL: DateTime queue failed"; exit 1 }
Write-Host "PASS"; exit 0
