# vybe-test: powershell/pipeline_tee_object_side_effects/tee_object_with_measure_object_downstream
$sideNums = $null
$meas = 1..10 | Tee-Object -Variable sideNums | Measure-Object -Sum -Average
if ($meas.Sum -ne 55 -or $meas.Average -ne 5.5 -or $sideNums.Count -ne 10) {
    Write-Host "FAIL: Tee-Object with downstream Measure-Object failed"
    exit 1
}
Write-Host "PASS"
exit 0
