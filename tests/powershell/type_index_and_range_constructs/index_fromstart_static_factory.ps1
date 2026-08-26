# vybe-test: powershell/type_index_and_range_constructs/index_fromstart_static_factory
$idx = [System.Index]::FromStart(5)
if ($idx.Value -ne 5 -or $idx.IsFromEnd) { Write-Host "FAIL: Index FromStart failed"; exit 1 }
Write-Host "PASS"; exit 0
