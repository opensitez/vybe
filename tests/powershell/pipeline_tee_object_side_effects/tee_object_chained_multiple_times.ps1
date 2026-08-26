# vybe-test: powershell/pipeline_tee_object_side_effects/tee_object_chained_multiple_times
$side1 = $null
$side2 = $null
$out = @(1, 2 | Tee-Object -Variable side1 | ForEach-Object { $_ + 10 } | Tee-Object -Variable side2 | ForEach-Object { $_ * 2 })
if ($side1[0] -ne 1 -or $side2[0] -ne 11 -or $out[0] -ne 22) {
    Write-Host "FAIL: Chained Tee-Object failed"
    exit 1
}
Write-Host "PASS"
exit 0
