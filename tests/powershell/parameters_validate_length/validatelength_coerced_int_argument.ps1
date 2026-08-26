# vybe-test: powershell/parameters_validate_length/validatelength_coerced_int_argument
function Set-IntString {
    param([ValidateLength(4, 6)][string]$Id)
    return $Id
}
$res = Set-IntString -Id 12345 # coerced to "12345" length 5
if ($res -ne "12345") {
    Write-Host "FAIL: ValidateLength with coerced integer string failed"
    exit 1
}
Write-Host "PASS"
exit 0
