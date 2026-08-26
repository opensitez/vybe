# vybe-test: powershell/collections_concurrent_queue/concurrent_queue_select_object_calculated
$cq = [System.Collections.Concurrent.ConcurrentQueue[int]]::new([int[]]@(1, 2, 3))
$res = @($cq | Select-Object @{ N = "Double"; E = { $_ * 2 } })
if ($res[0].Double -ne 2 -or $res[2].Double -ne 6) { Write-Host "FAIL: Select-Object calculated failed"; exit 1 }
Write-Host "PASS"; exit 0
