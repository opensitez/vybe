# vybe-test: powershell/pipeline_object_capture/capture_nested_subexpression
$res = $( $( 1..3 | Where-Object { $_ -gt 1 } ) | ForEach-Object { $_ * 5 } )
if ($res[0] -ne 10 -or $res[1] -ne 15) {
    Write-Host "FAIL: nested subexpression pipeline capture expected 10, 15"
    exit 1
}
Write-Host "PASS"
exit 0
