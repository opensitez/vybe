# vybe-test: powershell/language_null_coalescing_and_assignment/null_coalescing_with_false_is_not_null
$flag = $false
# Boolean $false is NOT null, so ?? must return $false
$res = $flag ?? $true
if ($res -ne $false) {
    Write-Host "FAIL: Boolean `$false should not be coalesced, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
