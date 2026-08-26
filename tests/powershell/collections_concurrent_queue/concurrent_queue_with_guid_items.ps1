# vybe-test: powershell/collections_concurrent_queue/concurrent_queue_with_guid_items
$cq = [System.Collections.Concurrent.ConcurrentQueue[guid]]::new()
$g = [guid]::NewGuid()
$cq.Enqueue($g)
[guid]$outGuid = [guid]::Empty
$ok = $cq.TryDequeue([ref]$outGuid)
if (-not $ok -or $outGuid -ne $g) { Write-Host "FAIL: Guid queue failed"; exit 1 }
Write-Host "PASS"; exit 0
