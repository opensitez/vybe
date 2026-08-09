# vybe-test: powershell/using_variable_scope/using_variable_boolean_flag
$flag = $true
$sb = { if ($using:flag) { "FLAG_ON" } else { "FLAG_OFF" } }
$res = &$sb
if ($res -ne "FLAG_ON") {
    Write-Host "FAIL: boolean using variable expected FLAG_ON, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
