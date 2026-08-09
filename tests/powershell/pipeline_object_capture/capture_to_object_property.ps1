# vybe-test: powershell/pipeline_object_capture/capture_to_object_property
$obj = [pscustomobject]@{ Data = $null }
$obj.Data = 1..3 | Where-Object { $_ -gt 1 }
if ($obj.Data.Count -ne 2 -or $obj.Data[0] -ne 2) {
    Write-Host "FAIL: pipeline capture to object property expected Data=@(2,3)"
    exit 1
}
Write-Host "PASS"
exit 0
