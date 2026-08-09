# vybe-test: powershell/out_variables/out_variable_hashtable_output
@{ K = 1 } | ForEach-Object { $_ } -OutVariable hCap | Out-Null
if ($hCap[0]["K"] -ne 1) {
    Write-Host "FAIL: hashtable OutVariable capture failed"
    exit 1
}
Write-Host "PASS"
exit 0
