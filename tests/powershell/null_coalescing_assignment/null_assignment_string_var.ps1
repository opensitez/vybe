# vybe-test: powershell/null_coalescing_assignment/null_assignment_string_var
$strVar = $null
$strVar ??= "DefaultString"
if ($strVar -ne "DefaultString") {
    Write-Host "FAIL: string variable ??= expected DefaultString, got $strVar"
    exit 1
}
Write-Host "PASS"
exit 0
