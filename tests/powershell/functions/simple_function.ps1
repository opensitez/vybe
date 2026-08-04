# vybe-test: powershell/functions/simple_function
function Get-Value {
    return 42
}
$result = Get-Value
if ($result -ne 42) {
    Write-Host "FAIL: expected 42, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
