# vybe-test: powershell/out_variables/out_variable_group_object
1..4 | Group-Object { $_ % 2 } -OutVariable grpCap | Out-Null
if ($grpCap.Count -ne 2) {
    Write-Host "FAIL: Group-Object OutVariable expected 2 group items"
    exit 1
}
Write-Host "PASS"
exit 0
