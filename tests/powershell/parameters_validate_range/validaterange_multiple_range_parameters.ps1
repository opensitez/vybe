# vybe-test: powershell/parameters_validate_range/validaterange_multiple_range_parameters
function Set-Window {
    param(
        [ValidateRange(100, 1920)][int]$Width,
        [ValidateRange(100, 1080)][int]$Height
    )
    return "$Width x $Height"
}
$res = Set-Window -Width 800 -Height 600
if ($res -ne "800 x 600") {
    Write-Host "FAIL: Multiple ValidateRange parameters failed"
    exit 1
}
Write-Host "PASS"
exit 0
