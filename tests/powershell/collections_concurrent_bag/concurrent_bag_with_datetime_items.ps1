# vybe-test: powershell/collections_concurrent_bag/concurrent_bag_with_datetime_items
$bag = [System.Collections.Concurrent.ConcurrentBag[datetime]]::new()
$dt = [datetime]::UtcNow
$bag.Add($dt)
[datetime]$outDt = [datetime]::MinValue
$ok = $bag.TryTake([ref]$outDt)
if (-not $ok -or $outDt -ne $dt) { Write-Host "FAIL: DateTime bag failed"; exit 1 }
Write-Host "PASS"; exit 0
