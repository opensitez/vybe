# vybe-test: powershell/pipeline_tee_object_side_effects/tee_object_in_middle_of_multi_stage_pipeline
$stage1 = $null
$final = @(1..5 | ForEach-Object { $_ * 2 } | Tee-Object -Variable stage1 | Where-Object { $_ -gt 4 })
if ($final.Length -ne 3 -or $stage1.Count -ne 5 -or $stage1[0] -ne 2) {
    Write-Host "FAIL: Tee-Object in middle of pipeline failed"
    exit 1
}
Write-Host "PASS"
exit 0
