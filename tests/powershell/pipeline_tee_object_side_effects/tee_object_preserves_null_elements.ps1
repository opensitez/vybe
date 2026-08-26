# vybe-test: powershell/pipeline_tee_object_side_effects/tee_object_preserves_null_elements
$arr = @("first", $null, "third")
$sideArr = $null
$res = @($arr | Tee-Object -Variable sideArr)
if ($res.Length -ne 3 -or $sideArr.Count -ne 3 -or $sideArr[1] -ne $null) {
    Write-Host "FAIL: Tee-Object null preservation failed"
    exit 1
}
Write-Host "PASS"
exit 0
