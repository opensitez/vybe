# vybe-test: powershell/language_ternary_conditional_operator/ternary_with_falsy_empty_array
$emptyArr = @()
$res = $emptyArr ? "NotEmpty" : "Empty"
if ($res -ne "Empty") {
    Write-Host "FAIL: Ternary with falsy empty array failed"
    exit 1
}
Write-Host "PASS"
exit 0
