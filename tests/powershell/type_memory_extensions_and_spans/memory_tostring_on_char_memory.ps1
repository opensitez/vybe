# vybe-test: powershell/type_memory_extensions_and_spans/memory_tostring_on_char_memory
$mem = [System.MemoryExtensions]::AsMemory("PowerShell")
if ($mem.ToString() -ne "PowerShell") { Write-Host "FAIL: ReadOnlyMemory.ToString failed"; exit 1 }
Write-Host "PASS"; exit 0
