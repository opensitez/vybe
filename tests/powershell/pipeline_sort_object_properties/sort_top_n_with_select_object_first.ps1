# vybe-test: powershell/pipeline_sort_object_properties/sort_top_n_with_select_object_first
$nums = @(10, 50, 20, 90, 80, 30)
$topTwo = @($nums | Sort-Object -Descending | Select-Object -First 2)
if ($topTwo.Length -ne 2 -or $topTwo[0] -ne 90 -or $topTwo[1] -ne 80) {
    Write-Host "FAIL: Sort-Object top N failed, got $($topTwo -join ',')"
    exit 1
}
Write-Host "PASS"
exit 0
