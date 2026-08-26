# vybe-test: powershell/operators/multiply_integers
$res = 6 * 7
if ($res -ne 42) {
    Write-Host "FAIL: Multiply integers failed"
    exit 1
}
Write-Host "PASS"
exit 0
