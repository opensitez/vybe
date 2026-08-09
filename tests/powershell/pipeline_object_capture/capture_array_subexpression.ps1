# vybe-test: powershell/pipeline_object_capture/capture_array_subexpression
$res = @(1..3 | ForEach-Object { $_ * 10 })
if ($res.Count -ne 3 -or $res[0] -ne 10 -or $res[2] -ne 30) {
    Write-Host "FAIL: capture array subexpression expected @(10, 20, 30)"
    exit 1
}
Write-Host "PASS"
exit 0
