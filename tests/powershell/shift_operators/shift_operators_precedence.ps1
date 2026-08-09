# vybe-test: powershell/shift_operators/shift_operators_precedence
$res = 2 + 3 -shl 1
if ($res -ne 10) {
    Write-Host "FAIL: 2 + 3 -shl 1 expected 10 (addition before shift), got $res"
    exit 1
}
Write-Host "PASS"
exit 0
