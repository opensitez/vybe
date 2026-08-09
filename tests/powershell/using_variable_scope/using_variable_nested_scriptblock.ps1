# vybe-test: powershell/using_variable_scope/using_variable_nested_scriptblock
$outerVal = 777
$outerSb = {
    $innerSb = { $using:outerVal }
    &$innerSb
}
$res = &$outerSb
if ($res -ne 777) {
    Write-Host "FAIL: nested scriptblock \$using: outerVal expected 777, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
