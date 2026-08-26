# vybe-test: powershell/scriptblock_closures/closure_pipeline_input
$factor = 5
$closureSb = { param($elem) $elem * $factor }.GetNewClosure()
$res = 1..3 | ForEach-Object { &$closureSb $_ }
if ($res[0] -ne 5 -or $res[2] -ne 15) {
    Write-Host "FAIL: closure in pipeline expected 5, 10, 15"
    exit 1
}
Write-Host "PASS"
exit 0
