# vybe-test: powershell/collections_concurrent_bag/concurrent_bag_trytake_empty_returns_false
$bag = [System.Collections.Concurrent.ConcurrentBag[int]]::new()
[int]$outVal = 0
$ok = $bag.TryTake([ref]$outVal)
if ($ok) { Write-Host "FAIL: TryTake empty should return false"; exit 1 }
Write-Host "PASS"; exit 0
