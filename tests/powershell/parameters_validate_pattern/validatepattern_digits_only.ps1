# vybe-test: powershell/parameters_validate_pattern/validatepattern_digits_only
function Set-Phone {
    param([ValidatePattern('^\d{10}$')][string]$Phone)
    return $Phone
}
$res = Set-Phone -Phone "1234567890"
if ($res -ne "1234567890") {
    Write-Host "FAIL: ValidatePattern digits only failed"
    exit 1
}
Write-Host "PASS"
exit 0
