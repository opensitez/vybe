# vybe-test: powershell/null_conditional/null_conditional_scriptblock_invoke
$sb = $null
$res = ${sb}?.Invoke()
if ($res -ne $null) {
    Write-Host "FAIL: null scriptblock invoke expected null, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
