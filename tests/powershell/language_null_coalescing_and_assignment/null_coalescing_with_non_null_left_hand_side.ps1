# vybe-test: powershell/language_null_coalescing_and_assignment/null_coalescing_with_non_null_left_hand_side
$left = "Primary"
$res = $left ?? "Fallback"
if ($res -ne "Primary") {
    Write-Host "FAIL: Null coalescing with non-null LHS failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
