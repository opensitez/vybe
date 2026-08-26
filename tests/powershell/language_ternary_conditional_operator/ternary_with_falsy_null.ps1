# vybe-test: powershell/language_ternary_conditional_operator/ternary_with_falsy_null
$val = $null
$res = $val ? "NotNull" : "IsNull"
if ($res -ne "IsNull") {
    Write-Host "FAIL: Ternary with falsy null failed"
    exit 1
}
Write-Host "PASS"
exit 0
