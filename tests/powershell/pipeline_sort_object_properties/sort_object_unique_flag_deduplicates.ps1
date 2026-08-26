# vybe-test: powershell/pipeline_sort_object_properties/sort_object_unique_flag_deduplicates
$nums = @(1, 2, 2, 3, 1, 4, 3)
$sorted = @($nums | Sort-Object -Unique)
if ($sorted.Length -ne 4 -or $sorted[0] -ne 1 -or $sorted[3] -ne 4) {
    Write-Host "FAIL: Sort-Object -Unique failed, got $($sorted -join ',')"
    exit 1
}
Write-Host "PASS"
exit 0
