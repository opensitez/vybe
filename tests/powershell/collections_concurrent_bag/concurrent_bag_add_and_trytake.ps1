# vybe-test: powershell/collections_concurrent_bag/concurrent_bag_add_and_trytake
$bag = [System.Collections.Concurrent.ConcurrentBag[int]]::new()
$bag.Add(10)
[int]$outVal = 0
$ok = $bag.TryTake([ref]$outVal)
if (-not $ok -or $outVal -ne 10 -or $bag.Count -ne 0) { Write-Host "FAIL: ConcurrentBag Add/TryTake failed"; exit 1 }
Write-Host "PASS"; exit 0
