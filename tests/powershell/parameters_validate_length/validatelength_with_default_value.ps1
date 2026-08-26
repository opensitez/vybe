# vybe-test: powershell/parameters_validate_length/validatelength_with_default_value
function Get-Handle {
    param([ValidateLength(3, 8)][string]$Handle = "user1")
    return $Handle
}
$res = Get-Handle
if ($res -ne "user1") {
    Write-Host "FAIL: ValidateLength default value failed"
    exit 1
}
Write-Host "PASS"
exit 0
