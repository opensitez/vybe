# vybe-test: powershell/type_index_and_range_constructs/range_startat_static_factory
$start = [System.Index]::FromStart(3)
$range = [System.Range]::StartAt($start)
if ($range.Start.Value -ne 3 -or -not $range.End.IsFromEnd) { Write-Host "FAIL: Range StartAt failed"; exit 1 }
Write-Host "PASS"; exit 0
