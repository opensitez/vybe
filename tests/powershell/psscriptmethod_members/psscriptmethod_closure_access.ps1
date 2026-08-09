# vybe-test: powershell/psscriptmethod_members/psscriptmethod_closure_access
$externalFactor = 10
$obj = [pscustomobject]@{ Base = 5 }
$obj | Add-Member -MemberType ScriptMethod -Name "Calc" -Value ({ $this.Base * $externalFactor }.GetClosure())
$res = $obj.Calc()
if ($res -ne 50) {
    Write-Host "FAIL: PSScriptMethod closure access expected 50, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
