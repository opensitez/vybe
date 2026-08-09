# vybe-test: powershell/null_conditional/null_conditional_chained_properties
$outer = [pscustomobject]@{ Inner = [pscustomobject]@{ Val = 99 } }
$res = ${outer}?.Inner?.Val
if ($res -ne 99) {
    Write-Host "FAIL: chained null-conditional properties expected 99, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
