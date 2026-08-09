# vybe-test: powershell/out_variables/out_variable_foreach_object
@("a", "b") | ForEach-Object { $_.ToUpper() } -OutVariable upperCap | Out-Null
if ($upperCap[0] -ne "A" -or $upperCap[1] -ne "B") {
    Write-Host "FAIL: ForEach-Object OutVariable expected A, B"
    exit 1
}
Write-Host "PASS"
exit 0
