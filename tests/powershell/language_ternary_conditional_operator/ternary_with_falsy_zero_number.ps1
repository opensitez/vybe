# vybe-test: powershell/language_ternary_conditional_operator/ternary_with_falsy_zero_number
$count = 0
$msg = $count ? "HasItems" : "Empty"
if ($msg -ne "Empty") {
    Write-Host "FAIL: Ternary with falsy zero number failed"
    exit 1
}
Write-Host "PASS"
exit 0
