# vybe-test: powershell/type_index_and_range_constructs/range_endat_static_factory
$end = [System.Index]::FromStart(7)
$range = [System.Range]::EndAt($end)
if ($range.Start.Value -ne 0 -or $range.End.Value -ne 7) { Write-Host "FAIL: Range EndAt failed"; exit 1 }
Write-Host "PASS"; exit 0
