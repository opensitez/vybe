# vybe-test: powershell/psvariable_objects/psvariable_new_variable_cmdlet
New-Variable -Name "NewVarCreated" -Value 789 -Description "Testing New-Variable"
if ($NewVarCreated -ne 789) {
    Write-Host "FAIL: New-Variable expected \$NewVarCreated=789"
    exit 1
}
Write-Host "PASS"
exit 0
