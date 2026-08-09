# vybe-test: powershell/psscriptmethod_members/psscriptmethod_add_member_cmdlet
$obj = [pscustomobject]@{}
Add-Member -InputObject $obj -MemberType ScriptMethod -Name "Action" -Value { "Done" }
$res = $obj.Action()
if ($res -ne "Done") {
    Write-Host "FAIL: Add-Member ScriptMethod expected Action()='Done', got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
