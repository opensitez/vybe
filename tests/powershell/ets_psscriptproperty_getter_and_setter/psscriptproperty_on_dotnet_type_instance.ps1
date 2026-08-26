# vybe-test: powershell/ets_psscriptproperty_getter_and_setter/psscriptproperty_on_dotnet_type_instance
$obj = [pscustomobject]@{ Tag = "NET" }
$obj | Add-Member -MemberType ScriptProperty -Name DoubleTag -Value { "$($this.Tag)$($this.Tag)" }
if ($obj.DoubleTag -ne "NETNET") {
    Write-Host "FAIL: PSScriptProperty on custom instance failed"
    exit 1
}
Write-Host "PASS"
exit 0
