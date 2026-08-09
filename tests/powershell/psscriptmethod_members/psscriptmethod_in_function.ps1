# vybe-test: powershell/psscriptmethod_members/psscriptmethod_in_function
function Attach-Method($o) {
    $o | Add-Member -MemberType ScriptMethod -Name "Hello" -Value { "HelloFromMethod" }
}
$obj = [pscustomobject]@{}
Attach-Method $obj
$res = $obj.Hello()
if ($res -ne "HelloFromMethod") {
    Write-Host "FAIL: function attached PSScriptMethod expected HelloFromMethod"
    exit 1
}
Write-Host "PASS"
exit 0
