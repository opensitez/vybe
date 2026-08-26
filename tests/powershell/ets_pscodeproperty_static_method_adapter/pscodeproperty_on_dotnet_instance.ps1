# vybe-test: powershell/ets_pscodeproperty_static_method_adapter/pscodeproperty_on_dotnet_instance
$obj = [pscustomobject]@{ Val = 100 }
$obj | Add-Member -MemberType ScriptProperty -Name Doubled -Value { $this.Val * 2 }
if ($obj.Doubled -ne 200) {
    Write-Host "FAIL: ETS property adapter failed"
    exit 1
}
Write-Host "PASS"
exit 0
