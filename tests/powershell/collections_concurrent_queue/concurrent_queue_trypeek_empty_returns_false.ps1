# vybe-test: powershell/collections_concurrent_queue/concurrent_queue_trypeek_empty_returns_false
$cq = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
[string]$outVal = ""
$ok = $cq.TryPeek([ref]$outVal)
if ($ok) { Write-Host "FAIL: TryPeek empty should return false"; exit 1 }
Write-Host "PASS"; exit 0
