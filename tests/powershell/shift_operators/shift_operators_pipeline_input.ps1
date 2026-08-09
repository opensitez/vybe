# vybe-test: powershell/shift_operators/shift_operators_pipeline_input
$res = 1..3 | ForEach-Object { 1 -shl $_ }
if ($res[0] -ne 2 -or $res[1] -ne 4 -or $res[2] -ne 8) {
    Write-Host "FAIL: pipeline shift left expected 2, 4, 8, got $($res -join ',')"
    exit 1
}
Write-Host "PASS"
exit 0
