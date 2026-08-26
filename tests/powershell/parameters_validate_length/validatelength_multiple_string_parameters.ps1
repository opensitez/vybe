# vybe-test: powershell/parameters_validate_length/validatelength_multiple_string_parameters
function Test-MultiLen {
    param(
        [ValidateLength(2, 4)][string]$A,
        [ValidateLength(2, 4)][string]$B
    )
    return "$A-$B"
}
$res = Test-MultiLen -A "ab" -B "cde"
if ($res -ne "ab-cde") {
    Write-Host "FAIL: Multiple ValidateLength parameters failed"
    exit 1
}
Write-Host "PASS"
exit 0
