# vybe-test: powershell/functions/advanced_function_cmdletbinding
function Test-Advanced {
    [CmdletBinding()]
    param($Value)
    return $Value * 2
}
$result = Test-Advanced -Value 21
if ($result -ne 42) {
    Write-Host "FAIL: expected 42, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
