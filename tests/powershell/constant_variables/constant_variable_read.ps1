# vybe-test: powershell/constant_variables/constant_variable_read
Set-Variable -Name "READ_ONLY_TEST" -Value 50 -Option Constant
$doubleVal = $READ_ONLY_TEST * 2
if ($doubleVal -ne 100) {
    Write-Host "FAIL: Constant variable reading in arithmetic expected 100, got $doubleVal"
    exit 1
}
Write-Host "PASS"
exit 0
