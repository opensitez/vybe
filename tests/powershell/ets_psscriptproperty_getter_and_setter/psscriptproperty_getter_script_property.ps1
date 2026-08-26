# vybe-test: powershell/ets_psscriptproperty_getter_and_setter/psscriptproperty_getter_script_property
$sb = { "Hello" }
$prop = [System.Management.Automation.PSScriptProperty]::new("Greeting", $sb)
if ($prop.GetterScript -ne $sb -or $prop.SetterScript -ne $null) {
    Write-Host "FAIL: GetterScript / SetterScript inspection failed"
    exit 1
}
Write-Host "PASS"
exit 0
