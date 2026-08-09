# vybe-test: powershell/psscriptmethod_members/psscriptmethod_array_return
$obj = [pscustomobject]@{}
$obj | Add-Member -MemberType ScriptMethod -Name "GetRange" -Value { 1..3 }
$res = $obj.GetRange()
if ($res.Count -ne 3 -or $res[2] -ne 3) {
    Write-Host "FAIL: PSScriptMethod array return expected Count 3, item 3"
    exit 1
}
Write-Host "PASS"
exit 0
