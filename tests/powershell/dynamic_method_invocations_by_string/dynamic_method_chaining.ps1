# vybe-test: powershell/dynamic_method_invocations_by_string/dynamic_method_chaining
$m1 = "Trim"
$m2 = "ToUpper"
$str = "   hello   "
$res = $str.$m1().$m2()
if ($res -ne "HELLO") {
    Write-Host "FAIL: Chained dynamic method invocations failed"
    exit 1
}
Write-Host "PASS"
exit 0
