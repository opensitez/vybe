# vybe-test: powershell/type_memory_extensions_and_spans/readonlymemory_char_slice_start
$str = "PrefixDataSuffix"
$mem = [System.MemoryExtensions]::AsMemory($str)
$sub = $mem.Slice(6)
if ($sub.ToString() -ne "DataSuffix") { Write-Host "FAIL: ReadOnlyMemory Slice start failed"; exit 1 }
Write-Host "PASS"; exit 0
