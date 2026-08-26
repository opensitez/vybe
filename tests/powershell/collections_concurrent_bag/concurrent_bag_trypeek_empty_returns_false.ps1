# vybe-test: powershell/collections_concurrent_bag/concurrent_bag_trypeek_empty_returns_false
$bag = [System.Collections.Concurrent.ConcurrentBag[string]]::new()
[string]$outVal = ""
$ok = $bag.TryPeek([ref]$outVal)
if ($ok) { Write-Host "FAIL: TryPeek empty should return false"; exit 1 }
Write-Host "PASS"; exit 0
