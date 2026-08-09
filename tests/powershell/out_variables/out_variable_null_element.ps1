# vybe-test: powershell/out_variables/out_variable_null_element
$null | ForEach-Object { $_ } -OutVariable nullCap | Out-Null
if ($nullCap.Count -ne 1 -or $nullCap[0] -ne $null) {
    Write-Host "FAIL: OutVariable null element capture expected 1 null item"
    exit 1
}
Write-Host "PASS"
exit 0
