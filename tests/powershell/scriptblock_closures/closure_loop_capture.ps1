# vybe-test: powershell/scriptblock_closures/closure_loop_capture
$closures = @()
foreach ($i in 1..3) {
    $closures += { $i }.GetNewClosure()
}
if ((&$closures[0]) -ne 1 -or (&$closures[2]) -ne 3) {
    Write-Host "FAIL: loop iteration closure capture expected 1, 2, 3"
    exit 1
}
Write-Host "PASS"
exit 0
