# vybe-test: powershell/psscriptmethod_members/psscriptmethod_overload_behavior
$obj = [pscustomobject]@{}
$obj | Add-Member -MemberType ScriptMethod -Name "Process" -Value {
    if ($args.Count -eq 0) { "NoArg" } else { "Arg:$($args[0])" }
}
$res1 = $obj.Process()
$res2 = $obj.Process("X")
if ($res1 -ne "NoArg" -or $res2 -ne "Arg:X") {
    Write-Host "FAIL: PSScriptMethod varargs overload expected NoArg, Arg:X"
    exit 1
}
Write-Host "PASS"
exit 0
