# vybe-test: powershell/collections_concurrent_stack/concurrent_stack_trypeek_element
$cs = [System.Collections.Concurrent.ConcurrentStack[string]]::new()
$cs.Push("Top")
[string]$outVal = ""
$ok = $cs.TryPeek([ref]$outVal)
if (-not $ok -or $outVal -ne "Top" -or $cs.Count -ne 1) { Write-Host "FAIL: TryPeek failed"; exit 1 }
Write-Host "PASS"; exit 0
