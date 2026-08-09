# vybe-test: powershell/pipeline_object_capture/capture_multiline_pipeline
$res = 1..10 |
    Where-Object { $_ -gt 5 } |
    ForEach-Object { $_ * 2 }
if ($res.Count -ne 5 -or $res[0] -ne 12 -or $res[4] -ne 20) {
    Write-Host "FAIL: multiline pipeline capture expected 12..20"
    exit 1
}
Write-Host "PASS"
exit 0
