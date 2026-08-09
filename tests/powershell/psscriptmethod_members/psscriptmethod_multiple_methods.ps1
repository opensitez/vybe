# vybe-test: powershell/psscriptmethod_members/psscriptmethod_multiple_methods
$obj = [pscustomobject]@{ Val = 100 }
$obj | Add-Member -MemberType ScriptMethod -Name "Inc" -Value { $this.Val += 10 }
$obj | Add-Member -MemberType ScriptMethod -Name "Dec" -Value { $this.Val -= 5 }
$obj.Inc()
$obj.Dec()
if ($obj.Val -ne 105) {
    Write-Host "FAIL: multiple PSScriptMethods execution expected Val=105, got $($obj.Val)"
    exit 1
}
Write-Host "PASS"
exit 0
