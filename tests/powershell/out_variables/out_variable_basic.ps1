# vybe-test: powershell/out_variables/out_variable_basic
$captured = $null
Get-Process -Id $PID -OutVariable captured | Out-Null
if ($captured -eq $null -or $captured[0].Id -ne $PID) {
    Write-Host "FAIL: -OutVariable basic capture failed"
    exit 1
}
Write-Host "PASS"
exit 0
