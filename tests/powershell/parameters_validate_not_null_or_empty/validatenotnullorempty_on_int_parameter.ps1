# vybe-test: powershell/parameters_validate_not_null_or_empty/validatenotnullorempty_on_int_parameter
function Set-NumVal {
    param([ValidateNotNullOrEmpty()][int]$Num)
    return $Num
}
$res = Set-NumVal -Num 0 # 0 is a valid non-null integer
if ($res -ne 0) {
    Write-Host "FAIL: Integer 0 failed ValidateNotNullOrEmpty"
    exit 1
}
Write-Host "PASS"
exit 0
