# vybe-test: powershell/pipeline_tee_object_side_effects/tee_object_with_select_object_first_short_circuit
$sideAll = $null
$res = @(1..100 | Tee-Object -Variable sideAll | Select-Object -First 3)
if ($res.Length -ne 3 -or $res[2] -ne 3) {
    Write-Host "FAIL: Tee-Object with Select-Object -First failed"
    exit 1
}
Write-Host "PASS"
exit 0
