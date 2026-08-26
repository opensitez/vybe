# vybe-test: powershell/type_memory_extensions_and_spans/memory_in_generic_list
$list = [System.Collections.Generic.List[System.ReadOnlyMemory[char]]]::new()
$list.Add([System.MemoryExtensions]::AsMemory("Item1"))
if ($list.Count -ne 1 -or $list[0].ToString() -ne "Item1") { Write-Host "FAIL: ReadOnlyMemory in List failed"; exit 1 }
Write-Host "PASS"; exit 0
