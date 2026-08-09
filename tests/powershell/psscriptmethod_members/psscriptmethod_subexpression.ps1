# vybe-test: powershell/psscriptmethod_members/psscriptmethod_subexpression
$obj = [pscustomobject]@{}
$obj | Add-Member -MemberType ScriptMethod -Name "GetName" -Value { "SubName" }
$msg = "Method: $( $obj.GetName() )"
if ($msg -ne "Method: SubName") {
    Write-Host "FAIL: PSScriptMethod subexpression expected 'Method: SubName', got '$msg'"
    exit 1
}
Write-Host "PASS"
exit 0
