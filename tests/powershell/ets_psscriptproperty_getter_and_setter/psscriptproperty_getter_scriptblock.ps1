# vybe-test: powershell/ets_psscriptproperty_getter_and_setter/psscriptproperty_getter_scriptblock
$obj = [pscustomobject]@{ First = "John"; Last = "Doe" }
$prop = [System.Management.Automation.PSScriptProperty]::new("FullName", { "$($this.First) $($this.Last)" })
$obj.PSObject.Properties.Add($prop)
if ($obj.FullName -ne "John Doe") {
    Write-Host "FAIL: PSScriptProperty getter failed, got '$($obj.FullName)'"
    exit 1
}
Write-Host "PASS"
exit 0
