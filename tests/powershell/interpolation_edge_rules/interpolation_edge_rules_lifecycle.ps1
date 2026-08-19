# vybe-test: powershell/interpolation_edge_rules/lifecycle
$outer = 'outside'
$innerCaptured = & {
    $outer = 'inside'
    "${outer}"
}
if ($innerCaptured -ne 'inside') {
    Write-Host "FAIL: inner scope interpolation expected inside"
    exit 1
}
if ($outer -ne 'outside') {
    Write-Host "FAIL: outer variable should remain outside after child scope"
    exit 1
}
Write-Host 'PASS'
exit 0
