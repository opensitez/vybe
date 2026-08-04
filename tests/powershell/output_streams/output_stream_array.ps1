# vybe-test: powershell/output_streams/output_stream_array
$results = @(Write-Output (1..5))
if ($results.Count -ne 5) {
    Write-Host "FAIL: expected 5 outputs"
    exit 1
}
Write-Host 'PASS'
exit 0
