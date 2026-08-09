# vybe-test: powershell/psscriptmethod_members/psscriptmethod_hashtable_return
$obj = [pscustomobject]@{}
$obj | Add-Member -MemberType ScriptMethod -Name "GetMap" -Value { @{ Status = "Ready" } }
$res = $obj.GetMap()
if ($res.Status -ne "Ready") {
    Write-Host "FAIL: PSScriptMethod hashtable return expected Status=Ready"
    exit 1
}
Write-Host "PASS"
exit 0
