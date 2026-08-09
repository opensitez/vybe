# vybe-test: powershell/null_conditional/null_conditional_index_null
$arr = $null
$res = ${arr}?[0]
if ($res -ne $null) {
    Write-Host "FAIL: null conditional index expected null, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
