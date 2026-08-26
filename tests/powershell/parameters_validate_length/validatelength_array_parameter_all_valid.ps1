# vybe-test: powershell/parameters_validate_length/validatelength_array_parameter_all_valid
function Set-Aliases {
    param([ValidateLength(2, 4)][string[]]$Aliases)
    return $Aliases.Length
}
$res = Set-Aliases -Aliases "abc", "xy", "test"
if ($res -ne 3) {
    Write-Host "FAIL: ValidateLength on array parameter failed"
    exit 1
}
Write-Host "PASS"
exit 0
