# vybe-test: powershell/pipeline_object_capture/capture_chained_pipeline_variables
$first = 1..3
$second = $first | ForEach-Object { $_ * 10 }
$third = $second | Where-Object { $_ -gt 15 }
if ($third.Count -ne 2 -or $third[0] -ne 20 -or $third[1] -ne 30) {
    Write-Host "FAIL: chained pipeline variables expected 20, 30"
    exit 1
}
Write-Host "PASS"
exit 0
