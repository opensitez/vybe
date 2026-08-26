# vybe-test: powershell/language_null_coalescing_and_assignment/null_coalescing_with_null_left_hand_side
$left = $null
$res = $left ?? "Fallback"
if ($res -ne "Fallback") {
    Write-Host "FAIL: Null coalescing with null LHS failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
