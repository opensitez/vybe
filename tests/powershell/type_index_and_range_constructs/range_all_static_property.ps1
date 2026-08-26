# vybe-test: powershell/type_index_and_range_constructs/range_all_static_property
$all = [System.Range]::All
if ($all.Start.Value -ne 0 -or $all.Start.IsFromEnd -or -not $all.End.IsFromEnd -or $all.End.Value -ne 0) {
    Write-Host "FAIL: Range All property failed"; exit 1
}
Write-Host "PASS"; exit 0
