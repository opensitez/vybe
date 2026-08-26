# vybe-test: powershell/parameters_validate_length/validatelength_zero_min_length_allows_empty_string
function Set-OptionalNote {
    param([ValidateLength(0, 50)][string]$Note)
    return "Note:$Note"
}
$res = Set-OptionalNote -Note ""
if ($res -ne "Note:") {
    Write-Host "FAIL: ValidateLength(0, 50) on empty string failed"
    exit 1
}
Write-Host "PASS"
exit 0
