# vybe-test: powershell/out_variables/out_variable_sort_object
@(3, 1, 2) | Sort-Object -OutVariable sortCap | Out-Null
if ($sortCap[0] -ne 1 -or $sortCap[2] -ne 3) {
    Write-Host "FAIL: Sort-Object OutVariable expected 1, 2, 3"
    exit 1
}
Write-Host "PASS"
exit 0
