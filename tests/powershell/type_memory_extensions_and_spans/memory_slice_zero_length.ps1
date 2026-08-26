# vybe-test: powershell/type_memory_extensions_and_spans/memory_slice_zero_length
$mem = [System.MemoryExtensions]::AsMemory("Hello")
$zeroSlice = $mem.Slice(2, 0)
if (-not $zeroSlice.IsEmpty -or $zeroSlice.Length -ne 0) { Write-Host "FAIL: Zero length slice failed"; exit 1 }
Write-Host "PASS"; exit 0
