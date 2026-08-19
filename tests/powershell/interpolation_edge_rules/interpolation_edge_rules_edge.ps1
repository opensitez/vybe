# vybe-test: powershell/interpolation_edge_rules/edge
$name = "A{B}"
if ("$name" -ne 'A{B}') {
    Write-Host "FAIL: braces in interpolated text should be literal when not expansion pattern"
    exit 1
}
Write-Host 'PASS'
exit 0
