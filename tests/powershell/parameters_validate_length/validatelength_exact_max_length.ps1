# vybe-test: powershell/parameters_validate_length/validatelength_exact_max_length
function Set-Code2 {
    param([ValidateLength(4, 8)][string]$Code)
    return $Code
}
$res = Set-Code2 -Code "ABCDEFGH" # length 8
if ($res -ne "ABCDEFGH") {
    Write-Host "FAIL: ValidateLength exact max length failed"
    exit 1
}
Write-Host "PASS"
exit 0
