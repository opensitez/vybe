# vybe-test: powershell/attributes/validate_not_null
function Test-Value {
    param(
        [ValidateNotNull()]
        $Value
    )
    return $Value
}
$result = Test-Value -Value 123
if ($result -ne 123) {
    Write-Host "FAIL: expected 123, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
