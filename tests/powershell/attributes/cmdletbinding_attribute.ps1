# vybe-test: powershell/attributes/cmdletbinding_attribute
function Test-Func {
    [CmdletBinding()]
    param()
    if ($PSCmdlet -eq $null) {
        Write-Host "FAIL: expected PSCmdlet to be present"
        exit 1
    }
    return "ok"
}
$result = Test-Func
if ($result -ne "ok") {
    Write-Host "FAIL: expected ok, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
