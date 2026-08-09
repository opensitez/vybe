# vybe-test: powershell/pipeline_object_capture/capture_chained_assignment
$a = $b = 1..3 | Select-Object -First 2
if ($a.Count -ne 2 -or $b.Count -ne 2 -or $a[1] -ne 2) {
    Write-Host "FAIL: chained pipeline assignment expected a=2 items, b=2 items"
    exit 1
}
Write-Host "PASS"
exit 0
