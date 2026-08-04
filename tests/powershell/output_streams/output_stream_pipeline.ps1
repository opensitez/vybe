# vybe-test: powershell/output_streams/output_stream_pipeline
$result = (1..3 | ForEach-Object { Write-Output $_ })
if ($result.Count -ne 3) {
    Write-Host "FAIL: expected pipeline output count 3"
    exit 1
}
Write-Host 'PASS'
exit 0
