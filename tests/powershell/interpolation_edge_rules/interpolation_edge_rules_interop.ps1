# vybe-test: powershell/interpolation_edge_rules/interop
$envVar = (Get-ChildItem Env:USERNAME).Value
if ([string]::IsNullOrEmpty($envVar)) {
    Write-Host 'FAIL: environment interop interpolation should return non-empty'
    exit 1
}
Write-Host 'PASS'
exit 0
