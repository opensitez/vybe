# vybe-test: powershell/ets_psscriptproperty_getter_and_setter/psscriptproperty_removal
$obj = [pscustomobject]@{ X = 1 }
$obj.PSObject.Properties.Add([System.Management.Automation.PSScriptProperty]::new("TempScript", { "temp" }))
$obj.PSObject.Properties.Remove("TempScript")
if ($obj.TempScript -ne $null) {
    Write-Host "FAIL: PSScriptProperty removal failed"
    exit 1
}
Write-Host "PASS"
exit 0
