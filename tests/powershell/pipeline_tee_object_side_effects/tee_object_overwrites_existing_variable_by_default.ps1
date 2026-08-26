# vybe-test: powershell/pipeline_tee_object_side_effects/tee_object_overwrites_existing_variable_by_default
$saved = "old_value"
$res = @(10, 20 | Tee-Object -Variable saved)
if ($saved.Count -ne 2 -or $saved[0] -ne 10) {
    Write-Host "FAIL: Tee-Object variable overwrite failed"
    exit 1
}
Write-Host "PASS"
exit 0
