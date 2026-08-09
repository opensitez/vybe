# vybe-test: powershell/null_conditional/null_conditional_method_non_null
$str = "hello"
$res = ${str}?.ToUpper()
if ($res -ne "HELLO") {
    Write-Host "FAIL: non-null conditional method expected 'HELLO', got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
