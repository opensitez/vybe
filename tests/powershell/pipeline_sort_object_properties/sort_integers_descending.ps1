# vybe-test: powershell/pipeline_sort_object_properties/sort_integers_descending
$nums = @(5, 1, 9, 3, 7)
$sorted = @($nums | Sort-Object -Descending)
if ($sorted[0] -ne 9 -or $sorted[1] -ne 7 -or $sorted[4] -ne 1) {
    Write-Host "FAIL: Sort-Object integers descending failed"
    exit 1
}
Write-Host "PASS"
exit 0
