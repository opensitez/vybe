# vybe-test: powershell/pipeline_sort_object_properties/sort_empty_pipeline_returns_empty
$sorted = @(@() | Sort-Object)
if ($sorted.Length -ne 0) {
    Write-Host "FAIL: Sort-Object on empty pipeline failed"
    exit 1
}
Write-Host "PASS"
exit 0
