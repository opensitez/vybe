# vybe-test: powershell/output_streams/output_stream_multiple
$result = @(Write-Output 1; Write-Output 2)
if ($result.Count -ne 2) {
    Write-Host "FAIL: expected 2 outputs"
    exit 1
}
Write-Host 'PASS'
exit 0
