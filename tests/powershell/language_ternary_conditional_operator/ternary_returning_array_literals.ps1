# vybe-test: powershell/language_ternary_conditional_operator/ternary_returning_array_literals
$multi = $true
$arr = $multi ? @(1, 2, 3) : @(1)
if ($arr.Length -ne 3 -or $arr[2] -ne 3) {
    Write-Host "FAIL: Ternary returning array literal failed"
    exit 1
}
Write-Host "PASS"
exit 0
