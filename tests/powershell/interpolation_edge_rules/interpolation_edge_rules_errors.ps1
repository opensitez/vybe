# vybe-test: powershell/interpolation_edge_rules/errors
$threw = $false
try {
    "$(1/0)" | Out-Null
} catch {
    $threw = $true
}
if (-not $threw) {
    Write-Host 'FAIL: interpolation division by zero should throw'
    exit 1
}
Write-Host 'PASS'
exit 0
