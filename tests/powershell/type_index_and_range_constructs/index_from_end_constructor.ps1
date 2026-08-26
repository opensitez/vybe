# vybe-test: powershell/type_index_and_range_constructs/index_from_end_constructor
$idx = [System.Index]::new(1, $true)
if ($idx.Value -ne 1 -or -not $idx.IsFromEnd) { Write-Host "FAIL: Index from end failed"; exit 1 }
Write-Host "PASS"; exit 0
