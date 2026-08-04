# vybe-test: powershell/functions/parameter_validation_validaterange
function Test-ValidateRange {
    param(
        [ValidateRange(1, 10)]
        [int]$Number
    )
    return $Number
}
$result = Test-ValidateRange -Number 5
if ($result -ne 5) {
    Write-Host "FAIL: expected 5, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
