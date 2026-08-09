# vybe-test: powershell/out_variables/out_variable_where_object
10..15 | Where-Object { $_ -gt 13 } -OutVariable whereCap | Out-Null
if ($whereCap.Count -ne 2 -or $whereCap[0] -ne 14) {
    Write-Host "FAIL: Where-Object OutVariable expected 14, 15"
    exit 1
}
Write-Host "PASS"
exit 0
