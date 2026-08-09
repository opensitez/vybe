# vybe-test: powershell/psscriptmethod_members/psscriptmethod_exception_propagation
$obj = [pscustomobject]@{}
$obj | Add-Member -MemberType ScriptMethod -Name "Fail" -Value { throw "MethodError" }
try {
    $obj.Fail()
    Write-Host "FAIL: PSScriptMethod exception expected throw"
    exit 1
} catch {
    Write-Host "PASS"
    exit 0
}
