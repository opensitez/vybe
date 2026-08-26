# vybe-test: powershell/type_memory_extensions_and_spans/readonlymemory_char_slice_start_and_length
$str = "PrefixDataSuffix"
$mem = [System.MemoryExtensions]::AsMemory($str)
$sub = $mem.Slice(6, 4)
if ($sub.ToString() -ne "Data") { Write-Host "FAIL: ReadOnlyMemory Slice start/len failed"; exit 1 }
Write-Host "PASS"; exit 0
