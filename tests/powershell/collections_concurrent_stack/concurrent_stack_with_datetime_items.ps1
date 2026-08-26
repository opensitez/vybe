# vybe-test: powershell/collections_concurrent_stack/concurrent_stack_with_datetime_items
$cs = [System.Collections.Concurrent.ConcurrentStack[datetime]]::new()
$dt = [datetime]::UtcNow
$cs.Push($dt)
[datetime]$outDt = [datetime]::MinValue
$ok = $cs.TryPop([ref]$outDt)
if (-not $ok -or $outDt -ne $dt) { Write-Host "FAIL: DateTime stack failed"; exit 1 }
Write-Host "PASS"; exit 0
