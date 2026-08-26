# vybe-test: powershell/type_index_and_range_constructs/index_fromend_static_factory
$idx = [System.Index]::FromEnd(2)
if ($idx.Value -ne 2 -or -not $idx.IsFromEnd) { Write-Host "FAIL: Index FromEnd failed"; exit 1 }
Write-Host "PASS"; exit 0
