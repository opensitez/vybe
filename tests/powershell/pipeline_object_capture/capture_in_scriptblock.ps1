# vybe-test: powershell/pipeline_object_capture/capture_in_scriptblock
$sb = { $inner = 10..12 | Measure-Object -Sum; $inner.Sum }
$res = &$sb
if ($res -ne 33) {
    Write-Host "FAIL: pipeline capture in scriptblock expected 33, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
