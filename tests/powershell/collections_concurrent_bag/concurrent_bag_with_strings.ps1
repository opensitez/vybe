# vybe-test: powershell/collections_concurrent_bag/concurrent_bag_with_strings
$bag = [System.Collections.Concurrent.ConcurrentBag[string]]::new()
$bag.Add("alpha"); $bag.Add("beta")
if ($bag.Count -ne 2) { Write-Host "FAIL: String bag count failed"; exit 1 }
Write-Host "PASS"; exit 0
