# vybe-test: powershell/interpolation_edge_rules/stream
$items = 1,2,3
$joined = "$($items -join ',')"
if ($joined -ne '1,2,3') {
    Write-Host "FAIL: interpolation over stream array should join via -join in expression, got '$joined'"
    exit 1
}
Write-Host 'PASS'
exit 0
