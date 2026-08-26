# vybe-test: powershell/ets_psscriptproperty_getter_and_setter/psscriptproperty_add_member_convenience_cmdlet
$obj = [pscustomobject]@{ Radius = 10 }
$obj | Add-Member -MemberType ScriptProperty -Name Circumference -Value { 2 * [math]::PI * $this.Radius }
if ([math]::Abs($obj.Circumference - 62.83185307) -gt 1e-4) {
    Write-Host "FAIL: Add-Member ScriptProperty failed, got $($obj.Circumference)"
    exit 1
}
Write-Host "PASS"
exit 0
