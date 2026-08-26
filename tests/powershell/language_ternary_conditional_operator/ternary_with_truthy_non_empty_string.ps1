# vybe-test: powershell/language_ternary_conditional_operator/ternary_with_truthy_non_empty_string
$str = "text"
$res = $str ? "HasText" : "EmptyText"
if ($res -ne "HasText") {
    Write-Host "FAIL: Ternary with truthy string failed"
    exit 1
}
Write-Host "PASS"
exit 0
