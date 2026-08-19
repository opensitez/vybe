# vybe-test: powershell/interpolation_edge_rules/nested
$data = [pscustomobject]@{ level = [pscustomobject]@{ value = 9 } }
if ("$($data.level.value)" -ne '9') {
    Write-Host "FAIL: nested interpolation expected 9"
    exit 1
}
Write-Host 'PASS'
exit 0
