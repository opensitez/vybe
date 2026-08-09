# vybe-test: powershell/null_conditional/null_conditional_index_non_null
$arr = @("A", "B", "C")
$res = ${arr}?[1]
if ($res -ne "B") {
    Write-Host "FAIL: non-null conditional index expected 'B', got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
