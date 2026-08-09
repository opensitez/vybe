# vybe-test: powershell/psvariable_objects/psvariable_constructor_basic
$v = [System.Management.Automation.PSVariable]::new("VarName", "VarVal")
if ($v.Name -ne "VarName" -or $v.Value -ne "VarVal") {
    Write-Host "FAIL: PSVariable constructor expected Name='VarName', Value='VarVal'"
    exit 1
}
Write-Host "PASS"
exit 0
