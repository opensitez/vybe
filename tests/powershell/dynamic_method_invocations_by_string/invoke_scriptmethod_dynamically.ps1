# vybe-test: powershell/dynamic_method_invocations_by_string/invoke_scriptmethod_dynamically
$obj = [pscustomobject]@{ Factor = 3 }
$obj | Add-Member -MemberType ScriptMethod -Name Multiply -Value { param($n) $this.Factor * $n }
$methodName = "Multiply"
$res = $obj.$methodName(10)
if ($res -ne 30) {
    Write-Host "FAIL: Dynamic ScriptMethod invocation failed, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
