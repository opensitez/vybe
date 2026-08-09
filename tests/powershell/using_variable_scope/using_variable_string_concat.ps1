# vybe-test: powershell/using_variable_scope/using_variable_string_concat
$prefix = "Hello"
$sb = { "$using:prefix World" }
$res = &$sb
if ($res -ne "Hello World") {
    Write-Host "FAIL: string interpolation with \$using:prefix expected 'Hello World', got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
