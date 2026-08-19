# vybe-test: powershell/interpolation_edge_rules/precedence
$x = 2
$y = 3
if ("$( $x + $y * 10 )" -ne '32') {
    Write-Host 'FAIL: arithmetic precedence inside interpolation incorrect'
    exit 1
}
Write-Host 'PASS'
exit 0
