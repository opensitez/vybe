# vybe-test: powershell/psvariable_objects/psvariable_set_variable_cmdlet
Set-Variable -Name "SetTestVar" -Value 456
if ($SetTestVar -ne 456) {
    Write-Host "FAIL: Set-Variable expected \$SetTestVar=456, got $SetTestVar"
    exit 1
}
Write-Host "PASS"
exit 0
