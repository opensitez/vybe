# vybe-test: powershell/out_variables/out_variable_measure_object
1..5 | Measure-Object -Sum -OutVariable measCap | Out-Null
if ($measCap[0].Sum -ne 15) {
    Write-Host "FAIL: Measure-Object OutVariable Sum expected 15, got $($measCap[0].Sum)"
    exit 1
}
Write-Host "PASS"
exit 0
