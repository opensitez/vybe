# vybe-test: powershell/interpolation_edge_rules/conversion
$num = 7
if ("$([string]$num)" -ne '7') {
    Write-Host "FAIL: explicit conversion interpolation expected 7"
    exit 1
}
Write-Host 'PASS'
exit 0
