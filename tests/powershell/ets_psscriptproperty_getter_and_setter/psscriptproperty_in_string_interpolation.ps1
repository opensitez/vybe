# vybe-test: powershell/ets_psscriptproperty_getter_and_setter/psscriptproperty_in_string_interpolation
$obj = [pscustomobject]@{ Code = "ALPHA" }
$obj.PSObject.Properties.Add([System.Management.Automation.PSScriptProperty]::new("Tag", { "TAG-$($this.Code)" }))
$str = "Result: $($obj.Tag)"
if ($str -ne "Result: TAG-ALPHA") {
    Write-Host "FAIL: PSScriptProperty string interpolation failed"
    exit 1
}
Write-Host "PASS"
exit 0
