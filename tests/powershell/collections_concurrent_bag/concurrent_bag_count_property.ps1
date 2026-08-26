# vybe-test: powershell/collections_concurrent_bag/concurrent_bag_count_property
$bag = [System.Collections.Concurrent.ConcurrentBag[int]]::new()
for ($i = 0; $i -lt 25; $i++) { $bag.Add($i) }
if ($bag.Count -ne 25) { Write-Host "FAIL: Count property failed"; exit 1 }
Write-Host "PASS"; exit 0
