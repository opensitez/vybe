# vybe-test: powershell/collections_concurrent_bag/concurrent_bag_is_empty_property
$bag = [System.Collections.Concurrent.ConcurrentBag[string]]::new()
if (-not $bag.IsEmpty) { Write-Host "FAIL: Initial IsEmpty should be true"; exit 1 }
$bag.Add("val")
if ($bag.IsEmpty) { Write-Host "FAIL: IsEmpty after Add should be false"; exit 1 }
Write-Host "PASS"; exit 0
