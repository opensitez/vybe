# vybe-test: powershell/pipeline_object_capture/capture_out_variable
$dummy = Write-Output "CapturedStr" -OutVariable outVar
if ($outVar[0] -ne "CapturedStr") {
    Write-Host "FAIL: OutVariable capture expected CapturedStr, got '$($outVar[0])'"
    exit 1
}
Write-Host "PASS"
exit 0
