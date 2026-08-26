# vybe-test: powershell/pipeline_sort_object_properties/sort_single_item_pipeline
$sorted = @(42 | Sort-Object)
if ($sorted.Length -ne 1 -or $sorted[0] -ne 42) {
    Write-Host "FAIL: Sort-Object single item failed"
    exit 1
}
Write-Host "PASS"
exit 0
