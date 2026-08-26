# vybe-test: powershell/language_null_coalescing_and_assignment/null_coalescing_with_zero_is_not_null
$num = 0
# Integer 0 is NOT null, so ?? must return 0
$res = $num ?? 100
if ($res -ne 0) {
    Write-Host "FAIL: Integer 0 should not be coalesced, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
