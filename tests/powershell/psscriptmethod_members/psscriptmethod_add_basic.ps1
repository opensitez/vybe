# vybe-test: powershell/psscriptmethod_members/psscriptmethod_add_basic
$obj = [pscustomobject]@{ Val = 5 }
$obj | Add-Member -MemberType ScriptMethod -Name "GetDouble" -Value { $this.Val * 2 }
$res = $obj.GetDouble()
if ($res -ne 10) {
    Write-Host "FAIL: PSScriptMethod GetDouble() expected 10, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
