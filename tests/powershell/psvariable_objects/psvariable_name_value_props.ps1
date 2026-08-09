# vybe-test: powershell/psvariable_objects/psvariable_name_value_props
$v = Get-Variable -Name "PID"
if ($v.Name -ne "PID" -or $v.Value -le 0) {
    Write-Host "FAIL: Get-Variable PID expected positive integer process ID"
    exit 1
}
Write-Host "PASS"
exit 0
