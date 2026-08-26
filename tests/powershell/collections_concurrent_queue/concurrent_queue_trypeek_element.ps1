# vybe-test: powershell/collections_concurrent_queue/concurrent_queue_trypeek_element
$cq = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
$cq.Enqueue("First")
[string]$outVal = ""
$ok = $cq.TryPeek([ref]$outVal)
if (-not $ok -or $outVal -ne "First" -or $cq.Count -ne 1) { Write-Host "FAIL: TryPeek failed"; exit 1 }
Write-Host "PASS"; exit 0
