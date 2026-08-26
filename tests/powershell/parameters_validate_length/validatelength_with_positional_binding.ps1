# vybe-test: powershell/parameters_validate_length/validatelength_with_positional_binding
function Set-Token {
    param([Parameter(Position=0)][ValidateLength(8, 16)][string]$Token)
    return $Token
}
$res = Set-Token "1234567890"
if ($res -ne "1234567890") {
    Write-Host "FAIL: Positional ValidateLength failed"
    exit 1
}
Write-Host "PASS"
exit 0
