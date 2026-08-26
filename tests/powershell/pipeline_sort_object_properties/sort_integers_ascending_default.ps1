# vybe-test: powershell/pipeline_sort_object_properties/sort_integers_ascending_default
$nums = @(5, 1, 9, 3, 7)
$sorted = @($nums | Sort-Object)
if ($sorted[0] -ne 1 -or $sorted[1] -ne 3 -or $sorted[4] -ne 9) {
    Write-Host "FAIL: Sort-Object integers ascending failed"
    exit 1
}
Write-Host "PASS"
exit 0
