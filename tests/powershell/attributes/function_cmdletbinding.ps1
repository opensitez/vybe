# vybe-test: powershell/attributes/function_cmdletbinding
function Test-Func {
    [CmdletBinding()]
    param()
    return "ok"
}
$result = Test-Func
if ($result -ne "ok") {
    Write-Host "FAIL: expected ok, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
