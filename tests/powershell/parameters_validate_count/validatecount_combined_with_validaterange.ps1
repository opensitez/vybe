# vybe-test: powershell/parameters_validate_count/validatecount_combined_with_validaterange
function Set-ValidatedNumbers {
    param(
        [ValidateCount(2, 3)]
        [ValidateRange(1, 100)]
        [int[]]$Numbers
    )
    return $Numbers.Length
}
$res = Set-ValidatedNumbers -Numbers 10, 20
if ($res -ne 2) {
    Write-Host "FAIL: ValidateCount combined with ValidateRange failed"
    exit 1
}
Write-Host "PASS"
exit 0
