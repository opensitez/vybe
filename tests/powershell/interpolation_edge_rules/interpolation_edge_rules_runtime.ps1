# vybe-test: powershell/interpolation_edge_rules/runtime
$now = Get-Date
if (("$($now.Year)" -match '^\d{4}$') -ne $true) {
    Write-Host "FAIL: runtime property interpolation invalid"
    exit 1
}
Write-Host 'PASS'
exit 0
