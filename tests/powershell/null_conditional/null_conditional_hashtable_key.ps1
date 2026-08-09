# vybe-test: powershell/null_conditional/null_conditional_hashtable_key
$dict = @{ K1 = "V1" }
$res = ${dict}?["K1"]
if ($res -ne "V1") {
    Write-Host "FAIL: hashtable null-conditional index expected 'V1', got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
