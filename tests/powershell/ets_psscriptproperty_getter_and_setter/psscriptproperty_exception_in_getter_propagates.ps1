# vybe-test: powershell/ets_psscriptproperty_getter_and_setter/psscriptproperty_exception_in_getter_propagates
$obj = [pscustomobject]@{ Base = 10 }
$obj | Add-Member -MemberType ScriptProperty -Name ThrowProp -Value { 10 } -SecondValue { throw "SetterFailed" }
$caught = $false
try {
    $obj.ThrowProp = 20
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: ScriptProperty setter exception propagation failed"
    exit 1
}
Write-Host "PASS"
exit 0
