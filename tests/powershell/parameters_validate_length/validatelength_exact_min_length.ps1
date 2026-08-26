# vybe-test: powershell/parameters_validate_length/validatelength_exact_min_length
function Set-Code {
    param([ValidateLength(4, 8)][string]$Code)
    return $Code
}
$res = Set-Code -Code "ABCD" # length 4
if ($res -ne "ABCD") {
    Write-Host "FAIL: ValidateLength exact min length failed"
    exit 1
}
Write-Host "PASS"
exit 0
