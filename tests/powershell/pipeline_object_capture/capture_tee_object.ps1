# vybe-test: powershell/pipeline_object_capture/capture_tee_object
$res = 1..3 | Tee-Object -Variable teeVar | ForEach-Object { $_ * 2 }
if ($teeVar.Count -ne 3 -or $teeVar[0] -ne 1 -or $res[0] -ne 2) {
    Write-Host "FAIL: Tee-Object capture variable expected teeVar=1..3, res=2,4,6"
    exit 1
}
Write-Host "PASS"
exit 0
