# vybe-test: powershell/pscmdlet_write_warning_record/pscmdlet_write_warning_direct_constructor_instantiation
# WarningRecord can be directly instantiated via its constructor and inspected for properties
$warningObj = [System.Management.Automation.WarningRecord]::new("Manually constructed warning record")

if ($warningObj.Message -ne "Manually constructed warning record") {
    Write-Host "FAIL: constructed WarningRecord message mismatch"
    exit 1
}

if ($warningObj.GetType().FullName -ne "System.Management.Automation.WarningRecord") {
    Write-Host "FAIL: type mismatch on constructed WarningRecord"
    exit 1
}

Write-Host "PASS"
exit 0
