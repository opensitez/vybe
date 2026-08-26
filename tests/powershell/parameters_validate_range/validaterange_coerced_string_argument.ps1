# vybe-test: powershell/parameters_validate_range/validaterange_coerced_string_argument
function Set-Timeout {
    param([ValidateRange(1, 60)][int]$Sec)
    return $Sec
}
$res = Set-Timeout -Sec "30" # string coerced to int
if ($res -ne 30) {
    Write-Host "FAIL: ValidateRange with coerced string failed"
    exit 1
}
Write-Host "PASS"
exit 0
