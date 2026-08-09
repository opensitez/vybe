# vybe-test: powershell/psscriptmethod_members/psscriptmethod_pass_thru
$obj = [pscustomobject]@{ Base = 1 }
$returned = $obj | Add-Member -MemberType ScriptMethod -Name "Execute" -Value { $this.Base * 10 } -PassThru
if ($returned.Execute() -ne 10) {
    Write-Host "FAIL: Add-Member ScriptMethod -PassThru expected Execute()=10"
    exit 1
}
Write-Host "PASS"
exit 0
