# vybe-test: powershell/attributes/validate_set_attribute
function Test-Choice {
    param(
        [ValidateSet('Red', 'Green', 'Blue')]
        [string]$Value
    )
    return $Value
}
$result = Test-Choice -Value 'Green'
if ($result -ne 'Green') {
    Write-Host "FAIL: expected Green, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
