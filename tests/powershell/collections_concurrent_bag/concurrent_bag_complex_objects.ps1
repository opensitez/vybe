# vybe-test: powershell/collections_concurrent_bag/concurrent_bag_complex_objects
$bag = [System.Collections.Concurrent.ConcurrentBag[pscustomobject]]::new()
$bag.Add([pscustomobject]@{ Id = 100 })
[pscustomobject]$outObj = $null
$ok = $bag.TryTake([ref]$outObj)
if (-not $ok -or $outObj.Id -ne 100) { Write-Host "FAIL: Complex object in ConcurrentBag failed"; exit 1 }
Write-Host "PASS"; exit 0
