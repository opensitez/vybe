# vybe-test: powershell/attributes/validate_range_attribute
function Test-Range {
    param(
        [ValidateRange(10, 20)]
        [int]$Value
    )
    return $Value
}
$result = Test-Range -Value 15
if ($result -ne 15) {
    Write-Host "FAIL: expected 15, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
