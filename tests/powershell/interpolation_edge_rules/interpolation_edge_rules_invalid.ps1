# vybe-test: powershell/interpolation_edge_rules/interpolation_edge_rules_invalid
$x = 10
$x += 5
$x *= 2
if ($x -eq 30) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
