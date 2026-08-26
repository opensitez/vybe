# vybe-test: powershell/collections_concurrent_bag/concurrent_bag_where_object_filter
$bag = [System.Collections.Concurrent.ConcurrentBag[string]]::new([string[]]@("cat", "dog", "elephant"))
$longWords = @($bag | Where-Object { $_.Length -gt 3 })
if ($longWords.Length -ne 1 -or $longWords[0] -ne "elephant") { Write-Host "FAIL: Where-Object on ConcurrentBag failed"; exit 1 }
Write-Host "PASS"; exit 0
