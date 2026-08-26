# vybe-test: powershell/pipeline_tee_object_side_effects/tee_object_with_empty_stream
$arr = @()
$out = @($arr | Tee-Object -Variable captured)
if ($out.Length -ne 0) {
    Write-Host "FAIL: Tee-Object with empty stream failed"
    exit 1
}
Write-Host "PASS"
exit 0
