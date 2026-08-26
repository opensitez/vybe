# vybe-test: powershell/type_memory_extensions_and_spans/string_trimstart_via_memory_extensions
$str = "  start"
$mem = [System.MemoryExtensions]::AsMemory($str)
$trimmed = [System.MemoryExtensions]::TrimStart($mem)
if ($trimmed.ToString() -ne "start") { Write-Host "FAIL: MemoryExtensions.TrimStart failed"; exit 1 }
Write-Host "PASS"; exit 0
