# vybe-test: powershell/type_memory_extensions_and_spans/memory_is_empty_property
$str = ""
$mem = [System.MemoryExtensions]::AsMemory($str)
if (-not $mem.IsEmpty -or $mem.Length -ne 0) { Write-Host "FAIL: Memory IsEmpty failed"; exit 1 }
Write-Host "PASS"; exit 0
