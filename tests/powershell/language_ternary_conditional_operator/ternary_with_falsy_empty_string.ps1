# vybe-test: powershell/language_ternary_conditional_operator/ternary_with_falsy_empty_string
$str = ""
$res = $str ? "HasText" : "EmptyText"
if ($res -ne "EmptyText") {
    Write-Host "FAIL: Ternary with falsy empty string failed"
    exit 1
}
Write-Host "PASS"
exit 0
