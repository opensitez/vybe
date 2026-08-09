# vybe-test: powershell/using_variable_scope/using_variable_array_indexing
$arr = @("Zero", "One", "Two")
$sb = { ($using:arr)[1] }
$res = &$sb
if ($res -ne "One") {
    Write-Host "FAIL: array indexing (\$using:arr)[1] expected 'One', got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
