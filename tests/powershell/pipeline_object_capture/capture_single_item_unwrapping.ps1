# vybe-test: powershell/pipeline_object_capture/capture_single_item_unwrapping
$res = 1..5 | Where-Object { $_ -eq 3 }
if ($res -is [array]) {
    Write-Host "FAIL: single item pipeline capture should unwrap to int, got array"
    exit 1
}
if ($res -ne 3) {
    Write-Host "FAIL: single item expected 3, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
