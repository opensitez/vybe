# vybe-test: powershell/pipeline_tee_object_side_effects/tee_object_with_boolean_stream
$bools = @($true, $false, $true)
$sideBools = $null
$res = @($bools | Tee-Object -Variable sideBools)
if ($sideBools.Count -ne 3 -or $sideBools[1] -ne $false) {
    Write-Host "FAIL: Tee-Object with boolean stream failed"
    exit 1
}
Write-Host "PASS"
exit 0
