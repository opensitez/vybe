# vybe-test: powershell/out_variables/out_variable_select_object
1..10 | Select-Object -First 2 -OutVariable selCap | Out-Null
if ($selCap.Count -ne 2 -or $selCap[1] -ne 2) {
    Write-Host "FAIL: Select-Object -First 2 OutVariable expected 1, 2"
    exit 1
}
Write-Host "PASS"
exit 0
