# vybe-test: powershell/pipeline_tee_object_side_effects/tee_object_preserves_object_types
[datetime]$dt = [datetime]::UtcNow
$sideDt = $null
$res = $dt | Tee-Object -Variable sideDt
if ($sideDt -isnot [datetime] -or $sideDt -ne $dt) {
    Write-Host "FAIL: Tee-Object type preservation failed"
    exit 1
}
Write-Host "PASS"
exit 0
