# vybe-test: powershell/console_output/write_output_simple
$output = 'hello'
Write-Output $output
if ($output -ne 'hello') {
    Write-Host "FAIL: expected hello"
    exit 1
}
Write-Host 'PASS'
exit 0
