# vybe-test: powershell/parameters_validate_script/validatescript_with_positional_arguments
function Set-PosScript {
    param([Parameter(Position=0)][ValidateScript({ $_ -gt 0 })][int]$A)
    return $A * 10
}
$res = Set-PosScript 5
if ($res -ne 50) {
    Write-Host "FAIL: Positional ValidateScript failed"
    exit 1
}
Write-Host "PASS"
exit 0
