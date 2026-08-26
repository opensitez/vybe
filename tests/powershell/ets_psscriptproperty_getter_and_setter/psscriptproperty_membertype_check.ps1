# vybe-test: powershell/ets_psscriptproperty_getter_and_setter/psscriptproperty_membertype_check
$obj = [pscustomobject]@{ A = 1 }
$obj.PSObject.Properties.Add([System.Management.Automation.PSScriptProperty]::new("Calc", { 42 }))
$m = $obj.PSObject.Properties["Calc"]
if ($m.MemberType -ne [System.Management.Automation.PSMemberTypes]::ScriptProperty) {
    Write-Host "FAIL: PSScriptProperty MemberType check failed"
    exit 1
}
Write-Host "PASS"
exit 0
