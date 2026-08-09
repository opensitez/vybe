# vybe-test: powershell/psscriptmethod_members/psscriptmethod_with_arguments
$calc = [pscustomobject]@{ Base = 10 }
$calc | Add-Member -MemberType ScriptMethod -Name "Add" -Value { param($n) $this.Base + $n }
$res = $calc.Add(32)
if ($res -ne 42) {
    Write-Host "FAIL: PSScriptMethod with argument expected 42, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
