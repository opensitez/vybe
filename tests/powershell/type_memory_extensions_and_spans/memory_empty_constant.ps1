# vybe-test: powershell/type_memory_extensions_and_spans/memory_empty_constant
$empty = [System.ReadOnlyMemory[char]]::Empty
if (-not $empty.IsEmpty -or $empty.Length -ne 0) { Write-Host "FAIL: ReadOnlyMemory.Empty failed"; exit 1 }
Write-Host "PASS"; exit 0
