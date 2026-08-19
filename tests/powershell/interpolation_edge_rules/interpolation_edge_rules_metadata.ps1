# vybe-test: powershell/interpolation_edge_rules/metadata
$meta = $PSVersionTable.PSVersion.ToString()
if ([string]::IsNullOrWhiteSpace($meta)) {
    Write-Host 'FAIL: metadata interpolation should produce a version string'
    exit 1
}
Write-Host 'PASS'
exit 0
