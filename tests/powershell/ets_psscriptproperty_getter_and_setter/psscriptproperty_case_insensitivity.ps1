# vybe-test: powershell/ets_psscriptproperty_getter_and_setter/psscriptproperty_case_insensitivity
$obj = [pscustomobject]@{ Val = 99 }
$obj.PSObject.Properties.Add([System.Management.Automation.PSScriptProperty]::new("DoubleVal", { $this.Val * 2 }))
if ($obj.doubleval -ne 198 -or $obj.DOUBLEVAL -ne 198) {
    Write-Host "FAIL: PSScriptProperty case-insensitivity failed"
    exit 1
}
Write-Host "PASS"
exit 0
