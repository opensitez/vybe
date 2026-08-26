# vybe-test: powershell/type_index_and_range_constructs/range_constructor_start_and_end
$start = [System.Index]::FromStart(2)
$end = [System.Index]::FromEnd(1)
$range = [System.Range]::new($start, $end)
if ($range.Start.Value -ne 2 -or $range.End.Value -ne 1) { Write-Host "FAIL: Range constructor failed"; exit 1 }
Write-Host "PASS"; exit 0
