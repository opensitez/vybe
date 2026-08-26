# vybe-test: powershell/type_index_and_range_constructs/index_from_start_constructor
$idx = [System.Index]::new(3, $false)
if ($idx.Value -ne 3 -or $idx.IsFromEnd) { Write-Host "FAIL: Index from start failed"; exit 1 }
Write-Host "PASS"; exit 0
