# vybe-test: powershell/collections_concurrent_bag/concurrent_bag_with_guid_items
$bag = [System.Collections.Concurrent.ConcurrentBag[guid]]::new()
$g = [guid]::NewGuid()
$bag.Add($g)
[guid]$outG = [guid]::Empty
$ok = $bag.TryTake([ref]$outG)
if (-not $ok -or $outG -ne $g) { Write-Host "FAIL: Guid bag failed"; exit 1 }
Write-Host "PASS"; exit 0
