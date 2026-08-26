# vybe-test: powershell/collections_concurrent_stack/concurrent_stack_with_guid_items
$cs = [System.Collections.Concurrent.ConcurrentStack[guid]]::new()
$g = [guid]::NewGuid()
$cs.Push($g)
[guid]$outG = [guid]::Empty
$ok = $cs.TryPop([ref]$outG)
if (-not $ok -or $outG -ne $g) { Write-Host "FAIL: Guid stack failed"; exit 1 }
Write-Host "PASS"; exit 0
