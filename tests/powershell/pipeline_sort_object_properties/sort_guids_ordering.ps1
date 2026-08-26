# vybe-test: powershell/pipeline_sort_object_properties/sort_guids_ordering
$g1 = [guid]::Parse("00000000-0000-0000-0000-000000000002")
$g2 = [guid]::Parse("00000000-0000-0000-0000-000000000001")
$sorted = @($g1, $g2 | Sort-Object)
if ($sorted[0] -ne $g2 -or $sorted[1] -ne $g1) {
    Write-Host "FAIL: Sort-Object GUIDs ordering failed"
    exit 1
}
Write-Host "PASS"
exit 0
