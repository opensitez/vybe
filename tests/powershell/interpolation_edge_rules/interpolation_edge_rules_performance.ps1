# vybe-test: powershell/interpolation_edge_rules/performance
$sum = 0
1..100 | ForEach-Object { $sum += "$_".Length }
if ($sum -ne 192) {
    Write-Host "FAIL: interpolation/iteration baseline should be deterministic, got $sum"
    exit 1
}
Write-Host 'PASS'
exit 0
