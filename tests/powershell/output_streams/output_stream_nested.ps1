# vybe-test: powershell/output_streams/output_stream_nested
$result = @(Write-Output (Write-Output 5))
if ($result[-1] -ne 5) {
    Write-Host "FAIL: expected nested output 5"
    exit 1
}
Write-Host 'PASS'
exit 0
