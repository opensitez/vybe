# vybe-test: powershell/parameters_validate_not_null_or_empty/non_empty_array_succeeds
function Set-ItemArr2 {
    param([ValidateNotNullOrEmpty()][string[]]$Items)
    return $Items.Length
}
$res = Set-ItemArr2 -Items "a", "b"
if ($res -ne 2) {
    Write-Host "FAIL: Non-empty array failed ValidateNotNullOrEmpty"
    exit 1
}
Write-Host "PASS"
exit 0
