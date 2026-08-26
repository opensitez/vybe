# vybe-test: powershell/type_memory_extensions_and_spans/string_trim_via_memory_extensions
$str = "  trimmed  "
$mem = [System.MemoryExtensions]::AsMemory($str)
$trimmed = [System.MemoryExtensions]::Trim($mem)
if ($trimmed.ToString() -ne "trimmed") { Write-Host "FAIL: MemoryExtensions.Trim failed, got '$($trimmed.ToString())'"; exit 1 }
Write-Host "PASS"; exit 0
