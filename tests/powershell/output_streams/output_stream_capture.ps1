# vybe-test: powershell/output_streams/output_stream_capture
$results = @(Write-Output 1,2,3)
if ($results.Count -ne 3) {
    Write-Host "FAIL: expected 3 outputs"
    exit 1
}
Write-Host 'PASS'
exit 0
