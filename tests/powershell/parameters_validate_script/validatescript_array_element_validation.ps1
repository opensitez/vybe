# vybe-test: powershell/parameters_validate_script/validatescript_array_element_validation
function Check-AllPositive {
    param([ValidateScript({ $_ -gt 0 })][int[]]$Nums)
    return $Nums.Length
}
$res = Check-AllPositive -Nums 1, 2, 3
if ($res -ne 3) {
    Write-Host "FAIL: ValidateScript array element validation failed"
    exit 1
}
Write-Host "PASS"
exit 0
