# vybe-test: powershell/pipeline_tee_object_side_effects/tee_object_with_sort_object_downstream
$sideUnsorted = $null
$sorted = @(5, 3, 1, 4, 2 | Tee-Object -Variable sideUnsorted | Sort-Object)
if ($sorted[0] -ne 1 -or $sorted[4] -ne 5 -or $sideUnsorted[0] -ne 5) {
    Write-Host "FAIL: Tee-Object with Sort-Object downstream failed"
    exit 1
}
Write-Host "PASS"
exit 0
