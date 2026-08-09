# vybe-test: powershell/pipeline_object_capture/capture_empty_pipeline
$res = 1..5 | Where-Object { $_ -gt 100 }
if ($res -ne $null) {
    Write-Host "FAIL: empty pipeline capture expected null, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
