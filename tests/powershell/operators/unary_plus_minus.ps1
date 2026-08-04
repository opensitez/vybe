# vybe-test: powershell/operators/unary_plus_minus
$x = 5
$neg = -$x
if ($neg -ne -5) {
    Write-Host "FAIL: expected -5, got $neg"
    exit 1
}
$pos = +$x
if ($pos -ne 5) {
    Write-Host "FAIL: expected 5, got $pos"
    exit 1
}
Write-Host "PASS"
exit 0
