# vybe-test: powershell/psscriptmethod_members/psscriptmethod_case_insensitivity
$obj = [pscustomobject]@{}
$obj | Add-Member -MemberType ScriptMethod -Name "MyMethod" -Value { "CaseOK" }
$res = $obj.mymethod()
if ($res -ne "CaseOK") {
    Write-Host "FAIL: case-insensitive PSScriptMethod invocation expected CaseOK, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
