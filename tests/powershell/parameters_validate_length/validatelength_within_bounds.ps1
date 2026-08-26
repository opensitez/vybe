# vybe-test: powershell/parameters_validate_length/validatelength_within_bounds
function Set-Password {
    param([ValidateLength(6, 12)][string]$Pwd)
    return "Valid:$Pwd"
}
$res = Set-Password -Pwd "secret12"
if ($res -ne "Valid:secret12") {
    Write-Host "FAIL: ValidateLength within bounds failed"
    exit 1
}
Write-Host "PASS"
exit 0
