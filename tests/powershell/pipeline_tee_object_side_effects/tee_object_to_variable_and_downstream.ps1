# vybe-test: powershell/pipeline_tee_object_side_effects/tee_object_to_variable_and_downstream
$captured = $null
$downstream = @(1, 2, 3 | Tee-Object -Variable captured | ForEach-Object { $_ * 10 })
if ($downstream[0] -ne 10 -or $downstream[2] -ne 30 -or $captured.Count -ne 3 -or $captured[0] -ne 1) {
    Write-Host "FAIL: Tee-Object to variable and downstream failed"
    exit 1
}
Write-Host "PASS"
exit 0
