# vybe-test: powershell/psscriptmethod_members/psscriptmethod_return_value
$obj = [pscustomobject]@{}
$obj | Add-Member -MemberType ScriptMethod -Name "GetText" -Value { return "ReturnedText" }
$res = $obj.GetText()
if ($res -ne "ReturnedText") {
    Write-Host "FAIL: PSScriptMethod explicit return expected ReturnedText, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
