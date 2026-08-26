# vybe-test: powershell/collections_concurrent_stack/concurrent_stack_where_object_filter
$cs = [System.Collections.Concurrent.ConcurrentStack[int]]::new([int[]]@(1..6))
$odds = @($cs | Where-Object { $_ % 2 -ne 0 })
if ($odds.Length -ne 3) { Write-Host "FAIL: Where-Object filter failed"; exit 1 }
Write-Host "PASS"; exit 0
