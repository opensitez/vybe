# vybe-test: powershell/language_ternary_conditional_operator/ternary_with_truthy_non_zero_number
$count = 5
$msg = $count ? "HasItems" : "Empty"
if ($msg -ne "HasItems") {
    Write-Host "FAIL: Ternary with truthy non-zero number failed"
    exit 1
}
Write-Host "PASS"
exit 0
