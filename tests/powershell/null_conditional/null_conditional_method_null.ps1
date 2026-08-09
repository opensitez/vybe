# vybe-test: powershell/null_conditional/null_conditional_method_null
$str = $null
$res = ${str}?.ToUpper()
if ($res -ne $null) {
    Write-Host "FAIL: null conditional method expected null, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
