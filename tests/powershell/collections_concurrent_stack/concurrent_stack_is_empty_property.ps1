# vybe-test: powershell/collections_concurrent_stack/concurrent_stack_is_empty_property
$cs = [System.Collections.Concurrent.ConcurrentStack[string]]::new()
if (-not $cs.IsEmpty) { Write-Host "FAIL: Initial IsEmpty should be true"; exit 1 }
$cs.Push("data")
if ($cs.IsEmpty) { Write-Host "FAIL: IsEmpty after Push should be false"; exit 1 }
Write-Host "PASS"; exit 0
