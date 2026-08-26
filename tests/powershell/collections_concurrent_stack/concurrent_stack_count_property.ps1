# vybe-test: powershell/collections_concurrent_stack/concurrent_stack_count_property
$cs = [System.Collections.Concurrent.ConcurrentStack[int]]::new()
for ($i = 0; $i -lt 30; $i++) { $cs.Push($i) }
if ($cs.Count -ne 30) { Write-Host "FAIL: Count property failed"; exit 1 }
Write-Host "PASS"; exit 0
