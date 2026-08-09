# vybe-test: powershell/pipeline_object_capture/capture_subexpression_multiple_statements
$res = $($a = 5; $b = 10; $a + $b)
if ($res -ne 15) {
    Write-Host "FAIL: capture subexpression multiple statements expected 15, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
