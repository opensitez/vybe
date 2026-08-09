# vybe-test: powershell/using_variable_scope/using_variable_type_preservation
$typedDouble = [double]9.99
$sb = { ($using:typedDouble).GetType().Name }
$res = &$sb
if ($res -ne "Double") {
    Write-Host "FAIL: type preservation expected Double, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
