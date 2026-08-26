# vybe-test: powershell/collections_concurrent_bag/concurrent_bag_trypeek_element
$bag = [System.Collections.Concurrent.ConcurrentBag[string]]::new()
$bag.Add("Item1")
[string]$outVal = ""
$ok = $bag.TryPeek([ref]$outVal)
if (-not $ok -or $outVal -ne "Item1" -or $bag.Count -ne 1) { Write-Host "FAIL: ConcurrentBag TryPeek failed"; exit 1 }
Write-Host "PASS"; exit 0
