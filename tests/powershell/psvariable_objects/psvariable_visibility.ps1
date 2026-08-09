# vybe-test: powershell/psvariable_objects/psvariable_visibility
$v = [System.Management.Automation.PSVariable]::new("VisVar", 100)
if ($v.Visibility -ne [System.Management.Automation.SessionStateEntryVisibility]::Public) {
    Write-Host "FAIL: PSVariable default Visibility expected Public"
    exit 1
}
Write-Host "PASS"
exit 0
