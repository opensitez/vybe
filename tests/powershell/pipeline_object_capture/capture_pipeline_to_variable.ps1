# vybe-test: powershell/pipeline_object_capture/capture_pipeline_to_variable
$res = 1..5 | Where-Object { $_ % 2 -eq 0 }
if ($res.Count -ne 2 -or $res[0] -ne 2 -or $res[1] -ne 4) {
    Write-Host "FAIL: capture pipeline to variable expected 2, 4"
    exit 1
}
Write-Host "PASS"
exit 0
