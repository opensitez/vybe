# vybe-test: powershell/using_variable_scope/using_variable_math_operation
$x = 9
$sb = { [Math]::Sqrt($using:x) }
$res = &$sb
if ($res -ne 3) {
    Write-Host "FAIL: [Math]::Sqrt(\$using:x) expected 3, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
