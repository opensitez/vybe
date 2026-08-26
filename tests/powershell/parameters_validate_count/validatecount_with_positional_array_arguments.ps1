# vybe-test: powershell/parameters_validate_count/validatecount_with_positional_array_arguments
function Set-PosArray {
    param([Parameter(Position=0)][ValidateCount(2, 3)][string[]]$ArgsList)
    return $ArgsList.Length
}
$res = Set-PosArray "one", "two"
if ($res -ne 2) {
    Write-Host "FAIL: Positional ValidateCount failed"
    exit 1
}
Write-Host "PASS"
exit 0
