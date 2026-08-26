# vybe-test: powershell/type_index_and_range_constructs/index_start_end_static_properties
$start = [System.Index]::Start
$end = [System.Index]::End
if ($start.Value -ne 0 -or $start.IsFromEnd -or -not $end.IsFromEnd -or $end.Value -ne 0) {
    Write-Host "FAIL: Index Start/End properties failed"; exit 1
}
Write-Host "PASS"; exit 0
