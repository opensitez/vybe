# vybe-test: powershell/type_memory_extensions_and_spans/string_trimend_via_memory_extensions
$str = "end  "
$mem = [System.MemoryExtensions]::AsMemory($str)
$trimmed = [System.MemoryExtensions]::TrimEnd($mem)
if ($trimmed.ToString() -ne "end") { Write-Host "FAIL: MemoryExtensions.TrimEnd failed"; exit 1 }
Write-Host "PASS"; exit 0
