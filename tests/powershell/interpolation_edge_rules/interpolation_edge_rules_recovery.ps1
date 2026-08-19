# vybe-test: powershell/interpolation_edge_rules/recovery
$val = if ($false) { 0 } else { 11 }
if ("$val" -ne '11') {
    Write-Host "FAIL: fallback branch interpolation expected 11"
    exit 1
}
Write-Host 'PASS'
exit 0
