# vybe-test: powershell/pipeline_sort_object_properties/sort_datetimes_chronological
$d1 = [datetime]::Parse("2026-12-01")
$d2 = [datetime]::Parse("2026-01-15")
$d3 = [datetime]::Parse("2026-06-30")
$sorted = @($d1, $d2, $d3 | Sort-Object)
if ($sorted[0].Month -ne 1 -or $sorted[1].Month -ne 6 -or $sorted[2].Month -ne 12) {
    Write-Host "FAIL: Sort-Object DateTime chronological failed"
    exit 1
}
Write-Host "PASS"
exit 0
