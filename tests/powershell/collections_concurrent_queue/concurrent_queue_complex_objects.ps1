# vybe-test: powershell/collections_concurrent_queue/concurrent_queue_complex_objects
$cq = [System.Collections.Concurrent.ConcurrentQueue[pscustomobject]]::new()
$cq.Enqueue([pscustomobject]@{ Val = 42 })
[pscustomobject]$outObj = $null
$ok = $cq.TryDequeue([ref]$outObj)
if (-not $ok -or $outObj.Val -ne 42) { Write-Host "FAIL: Complex object in ConcurrentQueue failed"; exit 1 }
Write-Host "PASS"; exit 0
