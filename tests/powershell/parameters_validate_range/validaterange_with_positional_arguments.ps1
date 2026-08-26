# vybe-test: powershell/parameters_validate_range/validaterange_with_positional_arguments
function Set-PosRange {
    param(
        [Parameter(Position=0)][ValidateRange(1, 5)][int]$A,
        [Parameter(Position=1)][ValidateRange(10, 50)][int]$B
    )
    return ($A + $B)
}
$res = Set-PosRange 3 20
if ($res -ne 23) {
    Write-Host "FAIL: Positional ValidateRange failed"
    exit 1
}
Write-Host "PASS"
exit 0
