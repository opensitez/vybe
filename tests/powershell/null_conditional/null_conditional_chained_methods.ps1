# vybe-test: powershell/null_conditional/null_conditional_chained_methods
$str = "  vybe  "
$res = ${str}?.Trim()?.ToUpper()
if ($res -ne "VYBE") {
    Write-Host "FAIL: chained null-conditional methods expected 'VYBE', got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
