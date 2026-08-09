# vybe-test: powershell/readonly_variables/readonly_variable_read
New-Variable -Name "READ_VAL" -Value 7 -Option ReadOnly
$res = $READ_VAL * 6
if ($res -ne 42) {
    Write-Host "FAIL: ReadOnly variable in arithmetic expected 42, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
